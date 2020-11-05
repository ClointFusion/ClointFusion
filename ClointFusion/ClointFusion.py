# Project Name: ClointFusion
# Project Description: A Python based RPA Automation Framework for Desktop GUI, Citrix, Web and basic Excel operations.

# Project Structure
# 1. All imports
# 2. All global variables
# 3. All function definitions
# 4. All test cases
# 5. All default services

# 1. All imports

import subprocess
import os
import sys
import platform
import urllib.request
import emoji
from pandas.core.algorithms import mode
from xlrd.formula import colname
from datetime import datetime
import pyautogui as pg
import time
import pandas as pd
import keyboard as kb
import PySimpleGUI as sg
import xlrd
import numpy
import openpyxl as op
from openpyxl import Workbook
from openpyxl import load_workbook
import datetime
import subprocess
from functools import lru_cache
import threading
from threading import Timer
import socket
from cv2 import cv2
import base64
import imutils
import clipboard
import re
from openpyxl import load_workbook
from openpyxl.styles import Font
from matplotlib.pyplot import axis
import plotly.express as px
from kaleido.scopes.plotly import PlotlyScope
import plotly.graph_objects as go
import zipcodes
import folium
from json import (load as jsonload, dump as jsondump)
from helium import *
from os import link
from selenium.webdriver import ChromeOptions
import dis
import texthero as hero
from texthero import preprocessing
from urllib.request import urlopen 
from hashlib import sha256
from PIL import Image
from wordcloud import WordCloud
from bs4 import BeautifulSoup
import requests
import watchdog.events
import watchdog.observers
from PyQt5 import QtWidgets, QtCore, QtGui
import tkinter as tk
from PIL import ImageGrab
from pathlib import Path
from pandas.core.common import flatten
from slack_webhook import Slack
import webbrowser 
import logging
import tempfile
from pif.utils import get_public_ip 
import pyautogui as pg
from email_validator import validate_email, EmailNotValidError
from skimage.metrics import structural_similarity
import warnings

os_name = str(platform.system()).lower()

#Windows OS specific packages
if os_name == "windows":
    from unicodedata import name
    import pygetwindow as gw    
sg.theme('Dark') # for PySimpleGUI FRONT END        

# 2. All global variables

base_dir = ""
config_folder_path = ""
log_path = ""
img_folder_path = ""
batch_file_path = ""
config_folder_path = ""
output_folder_path = ""
error_screen_shots_path = ""
status_log_excel_filepath = ""
bot_name = ""
current_working_dir = os.path.dirname(os.path.realpath(__file__)) #get cwd
temp_current_working_dir = tempfile.mkdtemp(prefix="cloint_",suffix="_fusion")
temp_current_working_dir = Path(temp_current_working_dir)
chrome_service = ""
browser_driver = ""
slack_webhook_url = "https://hooks.slack.com/services/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"
slack = Slack(url=slack_webhook_url)

cf_icon_file_path = Path(os.path.join(current_working_dir,"Cloint-ICON.ico"))
cf_logo_file_path = Path(os.path.join(current_working_dir,"Cloint-LOGO.PNG"))
ss_path_b = Path(os.path.join(config_folder_path,"my_screen_shot_before.png")) #before search
ss_path_a = Path(os.path.join(config_folder_path,"my_screen_shot_after.png")) #after search

enable_semi_automatic_mode = False
Browser_Service_Started = False
ai_screenshot = ""
ai_processes = []
helium_service_launched=False

# 3. All function definitions

#decorator to push a function to background using asyncio
def background(f):
    """
    Decorator function to push a function to background using asyncio
    """
    import asyncio
    try:
        from functools import wraps
        @wraps(f)
        def wrapped(*args, **kwargs):
            loop = asyncio.get_event_loop()
            if callable(f):
                return loop.run_in_executor(None, f, *args, **kwargs)
            else:
                raise TypeError('Task must be a callable')    
        return wrapped
    except Exception as ex:
        print("Task pushed to background = "+str(f) + str(ex))

def get_image_from_base64(imgFileName,imgBase64Str):
    """
    Function which converts the given Base64 string to an image and saves in given path

    Parameters:
        imgFileName  (str) : Image file name with png extension
        imgBase64Str (str) : Base64 string for conversion.
    """
    if not os.path.exists(imgFileName) :
        try:
            img_binary = base64.decodebytes(imgBase64Str)
            with open(imgFileName,"wb") as f:
                f.write(img_binary)
        except Exception as ex:
            print("Error in get_image_from_base64="+str(ex))

# @background
def _download_cloint_ico_png():    
    """
    Internal function to download ClointFusion ICON from GitHub
    """
    try:
        if not os.path.exists(cf_icon_file_path):
            urllib.request.urlretrieve('https://raw.githubusercontent.com/ClointFusion/Image_ICONS_GIFs/main/Cloint-ICON.ico',cf_icon_file_path)

        if not os.path.exists(cf_logo_file_path):
            urllib.request.urlretrieve('https://raw.githubusercontent.com/ClointFusion/Image_ICONS_GIFs/main/Cloint-LOGO.PNG',cf_logo_file_path)
    except Exception as ex:
        print("Error while downloading Cloint ICOn/LOGO = "+str(ex))

def show_emoji(strInput=""):
    """
    Function which prints Emojis

    Usage: 
    print(show_emoji('thumbsup'))
    print("OK",show_emoji('thumbsup'))
    Default: thumbsup
    """
    if not strInput:
        return(emoji.emojize(":{}:".format(str('thumbsup').lower()),use_aliases=True,variant="emoji_type"))
    else:
        return(emoji.emojize(":{}:".format(str(strInput).lower()),use_aliases=True,variant="emoji_type"))

def _load_missing_python_packages_windows():
    """
    Installs Windows OS specific python packages
    """       
    list_of_required_packages = ["pywin32","PyGetWindow"]

    try:
        reqs = subprocess.check_output([sys.executable, '-m', 'pip', 'list'])
        installed_packages = [r.decode().split('==')[0] for r in reqs.split()]
        missing_packages = ' '.join(list(set(list_of_required_packages)-set(installed_packages)))
        if missing_packages:
            print("{} package(s) are missing".format(missing_packages)) 
            
            os.system("{} -m pip install --upgrade pip".format(sys.executable))
            
            cmd = "pip install --upgrade {}".format(missing_packages)
            print(cmd)
            os.system(cmd) 

    except Exception as ex:
        print("Error in _load_missing_python_packages_windows="+str(ex))

def is_execution_required_today(function_name,execution_type="D",save_todays_date_month=False):
    """
    Function which ensures that a another function which calls this function is executed only once per day.
    Returns boolean True/False if another function to be executed today or not
    execution_type = D = Execute only once per day
    execution_type = M = Execute only once per month
    """
    if config_folder_path:
        last_updated_date_file = os.path.join(config_folder_path,function_name + ".txt")
    else:
        last_updated_date_file = os.path.join(current_working_dir,function_name + ".txt")

    last_updated_date_file = Path(last_updated_date_file)
    
    EXECUTE_NOW = False
    last_updated_on_date = ""
    
    if save_todays_date_month == False:
        try:    
            with open(last_updated_date_file, 'r') as f:
                last_updated_on_date = str(f.read())
        except:
            save_todays_date_month = True

    if save_todays_date_month:
        with open(last_updated_date_file, 'w',encoding="utf-8") as f:
            if execution_type == "D":
                last_updated_on_date = datetime.date.today().strftime('%d')
            elif execution_type == "M":
                last_updated_on_date = datetime.date.today().strftime('%m')
            f.write(str(last_updated_on_date))
            EXECUTE_NOW = True

    today_date_month = ""
    if execution_type == "D":
        today_date_month = str(datetime.date.today().strftime('%d'))
    elif execution_type == "M":
        today_date_month = str(datetime.date.today().strftime('%m'))

    if last_updated_on_date != today_date_month:
        EXECUTE_NOW = True

    return EXECUTE_NOW,last_updated_date_file

def _welcome_to_clointfusion():
    """
    Internal Function to display welcome message & push a notification to ClointFusion Slack
    """
    welcome_msg = "Welcome to ClointFusion, Made in India with " + show_emoji("red_heart")
    print(welcome_msg)

    hostname = socket.gethostname()
    public_ip = get_public_ip()
    my_ip = "Hostname : {}".format(hostname) + ", IP Address: " + str(socket.gethostbyname(hostname)) + " / " + str(public_ip)
    msg_status = str(my_ip) + ", BOT Name: " + bot_name + ", started at {}".format(time.strftime("%H:%M:%S on %Y-%m-%d"))
    slack.post(text=msg_status)

def _set_bot_name(strBotName=""):
    """
    Internal function
    If a botname is given, it will be used in the log file and in Task Scheduler
    we can also access the botname variable globally.

    Parameters :
        strBotName (str) : Name of the bot
    """
    global base_dir
    global bot_name

    if not strBotName: #if user has not given bot_name
        try:
            bot_name = current_working_dir[current_working_dir.rindex("\\") + 1 : ] #Assumption that user has given proper folder name and so taking it as BOT name
        except:
            bot_name = current_working_dir[current_working_dir.rindex("/") + 1 : ] #Assumption that user has given proper folder name and so taking it as BOT name

    else:
        strBotName = ''.join(e for e in strBotName if e.isalnum()) 
        bot_name = strBotName

    base_dir = str(base_dir) + "_" + bot_name
    base_dir = Path(base_dir)
    
def folder_create(strFolderPath=""):
    """
    while making leaf directory if any intermediate-level directory is missing,
    folder_create() method will create them all.

    Parameters:
        folderPath (str) : path to the folder where the folder is to be created.

    For example consider the following path:
    
    """
    try:
        if not strFolderPath:
            strFolderPath = gui_get_any_input_from_user('folder path to Create folder')

        if not os.path.exists(strFolderPath):
            os.makedirs(strFolderPath)
    except Exception as ex:
        print("Error in folder_create="+str(ex))

def _create_status_log_file(xtLogFilePath):
    """
    Internal Function to create Status Log File
    """
    try:
        if not os.path.exists(xtLogFilePath):
            df = pd.DataFrame({'Timestamp': [], 'Status':[]})
            writer = pd.ExcelWriter(xtLogFilePath)
            df.to_excel(writer, sheet_name='Sheet1', index=False)
            writer.save()
    except Exception as ex:
        print("Error in _create_status_log_file = " +str(ex))        
# @timeit
def _init_log_file():
    """
    Generates the log and saves it to the file in the given base directory. Internal function
    """
    global log_path
    global status_log_excel_filepath
    
    try:
        if bot_name:
            excelFileName = str(bot_name) + "-StatusLog.xlsx"
        else:
            excelFileName = "StatusLog.xlsx"

        folder_create(status_log_excel_filepath)
        
        status_log_excel_filepath = os.path.join(status_log_excel_filepath,excelFileName)
        status_log_excel_filepath = Path(status_log_excel_filepath)
        
        _create_status_log_file(status_log_excel_filepath)   

    except Exception as ex:
        print("ERROR in _init_log_file="+str(ex))

def _folder_read_text_file(txt_file_path=""):
    """
    Reads from a given text file and returns entire contents as a single list
    """
    try:
        with open(txt_file_path) as f:
            file_contents = f.read()
        return file_contents
    except:
        return None

def _folder_write_text_file(txt_file_path="",contents=""):
    """
    Writes given contents to a text file
    """
    try:
        
        f = open(txt_file_path,'w',encoding="utf-8")
        f.write(str(contents))
        f.close()
        
    except Exception as ex:
        print("Error in folder_write_text_file="+str(ex))

def _ask_user_semi_automatic_mode():
    """
    Ask user to 'Enable Semi Automatic Mode'
    """
    global enable_semi_automatic_mode
    values = []
    
    file_path = os.path.join(config_folder_path, 'Dont_Ask_Again.txt')
    file_path = Path(file_path)
    stored_do_not_ask_user_preference = _folder_read_text_file(file_path)

    file_path = os.path.join(config_folder_path, 'Semi_Automatic_Mode.txt')
    file_path = Path(file_path)
    enable_semi_automatic_mode = _folder_read_text_file(file_path)

    if enable_semi_automatic_mode:
        enable_semi_automatic_mode = enable_semi_automatic_mode[0]
    
    bot_config_path = os.path.join(config_folder_path,bot_name + ".xlsx")
    bot_config_path = Path(bot_config_path)

    if stored_do_not_ask_user_preference is None or str(stored_do_not_ask_user_preference[0]).lower() == 'false':

        layout = [[sg.Text('Do you want me to store GUI responses & use them next time when you run this BOT ?',text_color='orange',font='Courier 13')],
                [sg.Submit('Yes',bind_return_key=True,button_color=('white','green'),font='Courier 14'), sg.CloseButton('No', button_color=('white','firebrick'),font='Courier 14')],
                [sg.Checkbox('Do not ask me again', key='-DONT_ASK_AGAIN-',default=False, text_color='yellow',enable_events=True)],
                [sg.Text("To see this message again, goto 'Config_Files' folder of your BOT and change 'Dont_Ask_Again.txt' to False. \n Please find path here: {}".format(Path(os.path.join(config_folder_path, 'Dont_Ask_Again.txt'))),key='-DND-',visible=False,font='Courier 8')]]

        window = sg.Window('ClointFusion - Enable Semi Automatic Mode ?',layout,return_keyboard_events=True,use_default_focus=False,disable_close=False,element_justification='c',keep_on_top=True,finalize=True,icon=cf_icon_file_path)
        
        file_path = os.path.join(config_folder_path, 'Dont_Ask_Again.txt')
        file_path = Path(file_path)
        _folder_write_text_file(file_path,str(False))

        while True:
            event, values = window.read()
            if event == '-DONT_ASK_AGAIN-':
                stored_do_not_ask_user_preference = values['-DONT_ASK_AGAIN-']
                file_path = os.path.join(config_folder_path, 'Dont_Ask_Again.txt')
                file_path = Path(file_path)
                _folder_write_text_file(file_path,str(stored_do_not_ask_user_preference))

                if values['-DONT_ASK_AGAIN-']:
                    window.Element('-DND-').Update(visible=True)
                else:
                    window.Element('-DND-').Update(visible=False)
                    
            if event in (sg.WIN_CLOSED, 'No'): #ask me every time
                enable_semi_automatic_mode = False
                break
            elif event == 'Yes': #do not ask me again
                enable_semi_automatic_mode = True
                break
    
        window.close()

        if not os.path.exists(bot_config_path):
            df = pd.DataFrame({'SNO': [],'KEY': [], 'VALUE':[]})
            writer = pd.ExcelWriter(bot_config_path)
            df.to_excel(writer, sheet_name='Sheet1', index=False)
            writer.save()
            
        if enable_semi_automatic_mode:
            print("Semi Automatic Mode is ENABLED "+ show_emoji())
        else:
            print("Semi Automatic Mode is DISABLED "+ show_emoji())
        
        file_path = os.path.join(config_folder_path, 'Semi_Automatic_Mode.txt')
        file_path = Path(file_path)
        _folder_write_text_file(file_path,str(enable_semi_automatic_mode))

def timeit(method):
    """
    Decorator for computing time taken

    parameters:
        Method() name, by using @timeit just above the def: - defination of the function.

    returns:
        prints time take by the function 
    """
    def timed(*args, **kw):
        ts = time.time()
        result = method(*args, **kw)
        te = time.time()
        print('%r  %2.2f ms' % (method.__name__, (te - ts) * 1000))
        return result
    return timed

def read_semi_automatic_log(key):
    """
    Function to read a value from semi_automatic_log for a given key
    """
    try:
        if config_folder_path:
            bot_config_path = os.path.join(config_folder_path,bot_name + ".xlsx")
            bot_config_path = Path(bot_config_path)
        else:
            bot_config_path = os.path.join(current_working_dir,"First_Run.xlsx")
            bot_config_path = Path(bot_config_path)

            if not os.path.exists(bot_config_path):
                df = pd.DataFrame({'SNO': [],'KEY': [], 'VALUE':[]})
                writer = pd.ExcelWriter(bot_config_path)
                df.to_excel(writer, sheet_name='Sheet1', index=False)
                writer.save()

        df = pd.read_excel(bot_config_path)
        value = df[df['KEY'] == key]['VALUE'].to_list()
        value = str(value[0])
        return value

    except:
        return None

def _excel_if_value_exists(excel_path="",sheet_name='Sheet1',header=0,usecols="",value=""):
    """
    Check if a given value exists in given excel. Returns True / False
    """
    try:
        
        if usecols:
            df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header, usecols=usecols)
        else:
            df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header)
        
        if value in df.values:
            df = ''
            return True
        else:
            df = ''
            return False

    except Exception as ex:
        print("Error in _excel_if_value_exists="+str(ex))

def message_pop_up(strMsg="",delay=3):
    """
    Specified message will popup on the screen for a specified duration of time.

    Parameters:
        strMsg  (str) : message to popup.
        delay   (int) : duration of the popup.
    """
    try:
        # if not strMsg:
        #     strMsg = gui_get_any_input_from_user("pop-up message")
        sg.popup_no_wait(strMsg,title='ClointFusion',auto_close_duration=delay, auto_close=True, keep_on_top=True,background_color="white",text_color="black")#,icon=cloint_ico_logo_base64)
    except Exception as ex:
        print("Error in message_pop_up="+str(ex))

def update_semi_automatic_log(key, value):
    """
    Update semi automatic excel log 
    """
    try:
        if config_folder_path:
            bot_config_path = os.path.join(config_folder_path,bot_name + ".xlsx")
            
        else:
            bot_config_path = os.path.join(current_working_dir,"First_Run.xlsx")
        
        bot_config_path = Path(bot_config_path)

        if _excel_if_value_exists(bot_config_path,usecols=['KEY'],value=key):
            df = pd.read_excel(bot_config_path)
            row_index = df.index[df['KEY'] == key].tolist()[0]
            
            df.loc[row_index,'VALUE'] = value
            df.to_excel(bot_config_path,index=False)
        else:
            reader = pd.read_excel(bot_config_path)
            
            df = pd.DataFrame({'SNO': [len(reader)+1], 'KEY': [key], 'VALUE':[value]})
            writer = pd.ExcelWriter(bot_config_path, engine='openpyxl')
            writer.book = load_workbook(bot_config_path)
            writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
        
            df.to_excel(writer,index=False,header=False,startrow=len(reader)+1)
            writer.save()

    except Exception as ex:
        print("Error in update_semi_automatic_log="+str(ex))

def gui_get_any_file_from_user(msgForUser="the file : ",Extension_Without_Dot="*"):    
    """
    Generic function to accept file path from user using GUI. Returns the filepath value in string format.Default allows all files i.e *

    Default Text: "Please choose "
    """
    values = []
    try:
        oldValue = ""
        oldKey = msgForUser
        show_gui = False
        existing_value = read_semi_automatic_log(msgForUser)

        if existing_value is None:
            show_gui = True

        if str(enable_semi_automatic_mode).lower() == 'false' and existing_value:
            show_gui = True
            oldValue = existing_value
            
        if show_gui:
            layout = [[sg.Text("ClointFusion - Set Yourself Free for Better Work", font='Courier 16',text_color='orange')],
                [sg.Text('Please choose '),sg.Text(text=oldKey + " (ending with .{})".format(str(Extension_Without_Dot).lower()),font=('Courier 12'),text_color='yellow'),sg.Input(default_text=oldValue ,key='-FILE-', enable_events=True), sg.FileBrowse(file_types=((".{} File".format(Extension_Without_Dot), "*.{}".format(Extension_Without_Dot)),))],
                [sg.Submit('Done',button_color=('white','green'),bind_return_key=True),sg.CloseButton('Close',button_color=('white','firebrick'))]]

            window = sg.Window('ClointFusion',layout, return_keyboard_events=True,use_default_focus=True,disable_close=False,element_justification='c',keep_on_top=True, finalize=True,icon=cf_icon_file_path)

            while True:
                    event, values = window.read()
                    if event == sg.WIN_CLOSED or event == 'Close':
                        break
                    if event == 'Done':
                        if values['-FILE-']:
                            break
                        else:
                            message_pop_up("Please enter the required values")
                            # print("Please enter the values")
            window.close()

            if values and event == 'Done':
                values['-KEY-'] = msgForUser

                if str(values['-KEY-']) and str(values['-FILE-']):
                    update_semi_automatic_log(str(values['-KEY-']).strip(),str(values['-FILE-']).strip())
            
                if values is not None and str(values['-FILE-']):
                    return str(values['-FILE-']).strip()
            else:
                return None

        else:
            return str(existing_value)

    except Exception as ex:
        print("Error in gui_get_any_file_from_user="+str(ex))

def folder_read_text_file(txt_file_path=""):
    """
    Reads from a given text file and returns entire contents as a single list
    """
    try:
        if not txt_file_path:
            txt_file_path = gui_get_any_file_from_user('the text file to READ from',"txt")

        with open(txt_file_path) as f:
            file_contents = f.readlines()
        return file_contents
    except:
        return None

def folder_write_text_file(txt_file_path="",contents=""):
    """
    Writes given contents to a text file
    """
    try:
        
        if not txt_file_path:
            txt_file_path = gui_get_any_file_from_user('the text file to WRITE to',"txt")
            

        if not contents:
            contents = gui_get_any_input_from_user('text file contents')

        f = open(txt_file_path,'w',encoding="utf-8")
        f.writelines(str(contents))
        f.close()
        
    except Exception as ex:
        print("Error in folder_write_text_file="+str(ex))

def excel_get_all_sheet_names(excelFilePath=""):
    """
    Gives you all names of the sheets in the given excel sheet.

    Parameters:
        excelFilePath  (str) : Full path to the excel file with slashes.
    
    returns : 
        all the names of the excelsheets as a LIST.
    """
    try:
        if not excelFilePath:
            excelFilePath = gui_get_any_file_from_user("xlsx")

        xls = xlrd.open_workbook(excelFilePath, on_demand=True)
        return xls.sheet_names()
    except Exception as ex:
        print("Error in excel_get_all_sheet_names="+str(ex))
    
def message_counter_down_timer(start_value=5):
    """
    Function to show count-down timer. Default is 5 seconds.
    Ex: message_counter_down_timer()
    """
    CONTINUE = True
    layout = [[sg.Text('Calling ClointFusion Function in (seconds)',justification='c')],[sg.Text('',size=(10, 0),font=('Helvetica', 20),justification='c', key='text')],
            [sg.Exit(button_color=('white', 'firebrick4'), key='Cancel')]]

    window = sg.Window('ClointFusion - Countdown Timer', layout, no_titlebar=True, auto_size_buttons=False,keep_on_top=True, grab_anywhere=False, element_justification='c',element_padding=(0, 0),finalize=True,icon=cf_icon_file_path)

    current_value = start_value + 1

    while True:
        event, values = window.read(timeout=2)
        current_value = current_value - 1
        time.sleep(1)
            
        if current_value == 0:
            CONTINUE = True
            break
            
        if event in (sg.WIN_CLOSED, 'Cancel'):    
            CONTINUE = False  
            print("Action cancelled by user")
            break

        window['text'].update(value=current_value)

    window.close()
    return CONTINUE

def gui_get_consent_from_user(msgForUser="Continue ?"):    
    """
    Generic function to get consent from user using GUI. Returns the yes or no

    Default Text: "Do you want to "
    """
    values = []
    try:
        oldValue = ""
        oldKey = msgForUser
        show_gui = False
        existing_value = read_semi_automatic_log(msgForUser)

        if existing_value is None:
            show_gui = True

        if str(enable_semi_automatic_mode).lower() == 'false' and existing_value:
            show_gui = True
            oldValue = existing_value
            
        if show_gui:
            layout = [[sg.Text("ClointFusion - Set Yourself Free for Better Work", font='Courier 16',text_color='orange')],
                [sg.Text('Do you want to '),sg.Text(text=oldKey,font=('Courier 12'),text_color='yellow')],
                [sg.Submit('Yes',button_color=('white','green'),font=('Courier 14'),bind_return_key=True),sg.Submit('No',button_color=('white','firebrick'),font=('Courier 14'))]]

            window = sg.Window('ClointFusion',layout, return_keyboard_events=True,use_default_focus=True,disable_close=False,element_justification='c',keep_on_top=True, finalize=True,icon=cf_icon_file_path)

            while True:
                event, values = window.read()
                if event == 'No':
                    oldValue = 'No'
                    break
                if event == 'Yes':
                    oldValue = 'Yes'
                    break
                        
            window.close()
            values['-KEY-'] = msgForUser

            if str(values['-KEY-']):
                update_semi_automatic_log(str(values['-KEY-']).strip(),str(oldValue))
        
            return oldValue

        else:
            return str(existing_value)

    except Exception as ex:
        print("Error in gui_get_consent_from_user="+str(ex))

def gui_get_dropdownlist_values_from_user(msgForUser="",dropdown_list=[],multi_select=True): 
    """
    Generic function to accept one of the drop-down value from user using GUI. Returns all chosen values in list format.

    Default Text: "Please choose the item(s) from "
    """

    values = []
    dropdown_list = dropdown_list

    if dropdown_list:
        try:
            oldValue = []
            oldKey = msgForUser
            show_gui = False
            existing_value = read_semi_automatic_log(msgForUser)
            
            if existing_value is None:
                show_gui = True

            if str(enable_semi_automatic_mode).lower() == 'false' and existing_value:
                show_gui = True
                oldValue = existing_value
                
            if show_gui:
                if multi_select:
                    layout = [[sg.Text("ClointFusion - Set Yourself Free for Better Work", font='Courier 16', text_color='orange')],
                            [sg.Text('Please choose the item(s) from '),sg.Text(text=oldKey,font=('Courier 12'),text_color='yellow'),sg.Listbox(dropdown_list,size=(30, 5),key='-EXCELCOL-',default_values=oldValue,select_mode=sg.LISTBOX_SELECT_MODE_MULTIPLE,enable_events=True,change_submits=True)],#oldExcelCols
                            [sg.Submit('Done',button_color=('white','green'),bind_return_key=True),sg.CloseButton('Close',button_color=('white','firebrick'))]]

                else:
                    layout = [[sg.Text("ClointFusion - Set Yourself Free for Better Work", font='Courier 16', text_color='orange')],
                            [sg.Text('Please choose an item from '),sg.Text(text=oldKey,font=('Courier 12'),text_color='yellow'),sg.Listbox(dropdown_list,size=(30, 5),key='-EXCELCOL-',default_values=oldValue,select_mode=sg.LISTBOX_SELECT_MODE_SINGLE,enable_events=True,change_submits=True)],#oldExcelCols
                            [sg.Submit('Done',button_color=('white','green'),bind_return_key=True),sg.CloseButton('Close',button_color=('white','firebrick'))]]

                window = sg.Window('ClointFusion',layout, return_keyboard_events=True,use_default_focus=False,disable_close=False,element_justification='c',keep_on_top=True, finalize=True,icon=cf_icon_file_path)

                while True:                
                    event, values = window.read()
                    
                    if event is None or event == 'Cancel' or event == "Escape:27":
                        values = []
                        break

                    if event == 'Done':
                        if values and values['-EXCELCOL-']:
                            break
                        else:
                            message_pop_up("Please enter all the values")

                window.close()

                if values and event == 'Done':
                    values['-KEY-'] = msgForUser
                    
                    if str(values['-KEY-']) and str(values['-EXCELCOL-']):
                        update_semi_automatic_log(str(values['-KEY-']).strip(),str(values['-EXCELCOL-']).strip())

                    return values['-EXCELCOL-']
                else:
                    return oldValue
            
            else:
                return oldValue
                
        except Exception as ex:
            print("Error in gui_get_dropdownlist_values_from_user="+str(ex))
    else:
        print('gui_get_dropdownlist_values_from_user - List is empty')

def gui_get_excel_sheet_header_from_user(msgForUser=""): 
    """
    Generic function to accept excel path, sheet name and header from user using GUI. Returns all these values in disctionary format.

    Default Text: "Please choose the excel "
    """
    values = []
    sheet_namesLst = []
    try:
        oldValue = "" + "," + "Sheet1" + "," + "0"
        oldKey = msgForUser
        show_gui = False
        existing_value = read_semi_automatic_log(msgForUser)
        
        if existing_value is None:
            show_gui = True

        if str(enable_semi_automatic_mode).lower() == 'false' and existing_value:
            show_gui = True
            oldValue = existing_value
            
        if show_gui:
            oldFilePath, oldSheet , oldHeader = str(oldValue).split(",")
    
            layout = [[sg.Text("ClointFusion - Set Yourself Free for Better Work", font='Courier 16', text_color='orange')],
                    [sg.Text('Please choose the excel '),sg.Text(text=oldKey,font=('Courier 12'),text_color='yellow'),sg.Input(default_text=oldFilePath,key="-FILEPATH-",enable_events=True,change_submits=True), sg.FileBrowse(file_types=(("Excel File", "*.xls"),("Excel File", "*.xlsx")))], 
                    [sg.Text('Sheet Name'), sg.Combo(sheet_namesLst,default_value=oldSheet,size=(20, 0),key="-SHEET-",enable_events=True)], 
                    [sg.Text('Choose the header row'),sg.Spin(values=('0', '1', '2', '3', '4', '5'),initial_value=int(oldHeader),key="-HEADER-",enable_events=True,change_submits=True)],
                    # [sg.Checkbox('Use this excel file for all the excel related operations of this BOT',enable_events=True, key='-USE_THIS_EXCEL-',default=old_Use_This_excel, text_color='yellow')],
                    [sg.Submit('Done',button_color=('white','green'),bind_return_key=True),sg.CloseButton('Close',button_color=('white','firebrick'))]]
        
            window = sg.Window('ClointFusion',layout, return_keyboard_events=True,use_default_focus=False,disable_close=False,element_justification='c',keep_on_top=True, finalize=True,icon=cf_icon_file_path)

            while True:
                if oldFilePath: 
                    sheet_namesLst = excel_get_all_sheet_names(oldFilePath)
                    window['-SHEET-'].update(values=sheet_namesLst)   
                
                event, values = window.read()
                
                if event is None or event == 'Cancel' or event == "Escape:27":
                    values = []
                    break

                if event == 'Done':
                    if values and values['-FILEPATH-'] and values['-SHEET-']:
                        break
                    else:
                        message_pop_up("Please enter all the values")

                if event == '-FILEPATH-':
                    sheet_namesLst = excel_get_all_sheet_names(values['-FILEPATH-'])
                    window['-SHEET-'].update(values=sheet_namesLst)   
                    window.refresh()
                    oldFilePath = ""

                    if len(sheet_namesLst) >= 1:
                        window['-SHEET-'].update(value=sheet_namesLst[0]) 

                if event == '-SHEET-':
                    window['-SHEET-'].update(value=values['-SHEET-'])

            window.close()
            values['-KEY-'] = msgForUser
            
            concatenated_value = values['-FILEPATH-'] + "," +  values ['-SHEET-'] + "," + values['-HEADER-']
            
            if str(values['-KEY-']) and concatenated_value:
                update_semi_automatic_log(str(values['-KEY-']).strip(),str(concatenated_value))

            return values['-FILEPATH-'] , values ['-SHEET-'] , int(values['-HEADER-'])
        
        else:
            oldFilePath, oldSheet , oldHeader = str(existing_value).split(",")
            return oldFilePath, oldSheet , int(oldHeader)
            
    except Exception as ex:
        print("Error in gui_get_excel_sheet_header_from_user="+str(ex))
    
def gui_get_folder_path_from_user(msgForUser="the folder : "):    
    """
    Generic function to accept folder path from user using GUI. Returns the folderpath value in string format.

    Default text: "Please choose "
    """
    values = []
    try:
        oldValue = ""
        oldKey = msgForUser
        show_gui = False
        existing_value = read_semi_automatic_log(msgForUser)

        if existing_value is None:
            show_gui = True

        if str(enable_semi_automatic_mode).lower() == 'false' and existing_value:
            show_gui = True
            oldValue = existing_value
            
        if show_gui:
            layout = [[sg.Text("ClointFusion - Set Yourself Free for Better Work", font='Courier 16',text_color='orange')],
                [sg.Text('Please choose '),sg.Text(text=oldKey,font=('Courier 12'),text_color='yellow'),sg.Input(default_text=oldValue ,key='-FOLDER-', enable_events=True), sg.FolderBrowse()],
                [sg.Submit('Done',button_color=('white','green'),bind_return_key=True),sg.CloseButton('Close',button_color=('white','firebrick'))]]

            window = sg.Window('ClointFusion',layout, return_keyboard_events=True,use_default_focus=True,disable_close=False,element_justification='c',keep_on_top=True,finalize=True,icon=cf_icon_file_path)

            while True:
                event, values = window.read()

                if event == sg.WIN_CLOSED or event == 'Close':
                    break
                if event == 'Done':
                    if values and values['-FOLDER-']:
                        break
                    else:
                        message_pop_up("Please enter the required values")
            
            window.close()

            if values and event == 'Done':
                values['-KEY-'] = msgForUser

                if str(values['-KEY-']) and str(values['-FOLDER-']):
                    update_semi_automatic_log(str(values['-KEY-']).strip(),str(values['-FOLDER-']).strip())
            
                if values is not None:
                    return str(values['-FOLDER-']).strip()
            else:
                return None

        else:
            return str(existing_value)

    except Exception as ex:
        print("Error in gui_get_folder_path_from_user="+str(ex))

def gui_get_any_input_from_user(msgForUser="the value : ",password=False,mandatory_field=True):    
    """
    Generic function to accept any input (text / numeric) from user using GUI. Returns the value in string format.
    Please use unique message (key) for each value.

    Default Text: "Please enter "
    """
    values = []
    try:
        oldValue = ""
        oldKey = msgForUser
        show_gui = False
        existing_value = read_semi_automatic_log(msgForUser)

        if existing_value == "nan":
            existing_value = None
            
        if existing_value is None:
            show_gui = True

        if str(enable_semi_automatic_mode).lower() == 'false' and existing_value:
            show_gui = True
            oldValue = existing_value
            
        layout = ""
        if show_gui:
            if password:
                layout = [[sg.Text("ClointFusion - Set Yourself Free for Better Work", font='Courier 16',text_color='orange')],
                    [sg.Text('Please enter '),sg.Text(text=oldKey,font=('Courier 12'),text_color='yellow'),sg.Input(default_text=oldValue,key='-VALUE-', justification='c',password_char='*')],
                    [sg.Text('This field is mandatory',text_color='red')],
                    [sg.Submit('Done',button_color=('white','green'),bind_return_key=True),sg.CloseButton('Close',button_color=('white','firebrick'))]]

            elif not password and mandatory_field:
                layout = [[sg.Text("ClointFusion - Set Yourself Free for Better Work", font='Courier 16',text_color='orange')],
                    [sg.Text('Please enter '),sg.Text(text=oldKey,font=('Courier 12'),text_color='yellow'),sg.Input(default_text=oldValue,key='-VALUE-', justification='c')],
                    [sg.Text('This field is mandatory',text_color='red')],
                    [sg.Submit('Done',button_color=('white','green'),bind_return_key=True),sg.CloseButton('Close',button_color=('white','firebrick'))]]

            elif not password and not mandatory_field:
                layout = [[sg.Text("ClointFusion - Set Yourself Free for Better Work", font='Courier 16',text_color='orange')],
                    [sg.Text('Please enter '),sg.Text(text=oldKey,font=('Courier 12'),text_color='yellow'),sg.Input(default_text=oldValue,key='-VALUE-', justification='c')],
                    [sg.Submit('Done',button_color=('white','green'),bind_return_key=True),sg.CloseButton('Close',button_color=('white','firebrick'))]]

            window = sg.Window('ClointFusion',layout, return_keyboard_events=True,use_default_focus=True,disable_close=True,element_justification='c',keep_on_top=True,finalize=True,icon=cf_icon_file_path)

            while True:
                
                event, values = window.read()

                if event == sg.WIN_CLOSED or event == 'Close':
                    
                    if oldValue or (values and values['-VALUE-']):
                        break

                    else:
                        if mandatory_field:
                            message_pop_up("Its a mandatory field !.. Cannot proceed, exiting now..")
                            print("Exiting ClointFusion, as Mandatory field is missing")
                            sys.exit(0)
                        else:
                            print("Mandatory field is missing, continuing with None/Empty value")
                            break
                
                if event == 'Done':
                    if values['-VALUE-']:
                        break
                    else:
                        if mandatory_field:
                            message_pop_up("This value is required. Please enter the value..")
                        else:
                            break
            
            window.close()

            if values and event == 'Done':
                values['-KEY-'] = msgForUser
            
            if values is not None and str(values['-KEY-']) and str(values['-VALUE-']):
                update_semi_automatic_log(str(values['-KEY-']).strip(),str(values['-VALUE-']).strip())

            if values is not None and str(values['-VALUE-']):
                return str(values['-VALUE-']).strip()
            else:
                return None
        
        else:
            return str(existing_value)

    except Exception as ex:
        print("Error in gui_get_any_input_from_user="+str(ex))

def excel_get_all_header_columns(excel_path="",sheet_name="Sheet1",header=0):
    """
    Gives you all column header names of the given excel sheet.
    """
    col_lst = []
    try:
        if not excel_path:
            excel_path,sheet_name,header = gui_get_excel_sheet_header_from_user('to all header columns as a list')

        col_lst = pd.read_excel(excel_path,sheet_name=sheet_name,header=header,nrows=1,dtype=str).columns.tolist()
        return col_lst
    except Exception as ex:
        print("Error in excel_get_all_header_columns="+str(ex))

def _extract_filename_from_filepath(strFilePath=""):
    """
    Function which extracts file name from the given filepath
    """
    if strFilePath:
        try:
            strFileName = Path(strFilePath).name
            strFileName = str(strFileName).split(".")[0]
            # strFileName = strFilePath[strFilePath.rindex("\\") + 1 : ]
            # strFileName = strFileName.split(".")[0]
            return strFileName
        except Exception as ex:
            print("Error in _extract_filename_from_filepath="+str(ex))


    else:
        print("Please enter the value="+str(strFilePath))    
    
def string_remove_special_characters(inputStr=""):
    """
    Removes all the special character.

    Parameters:
        inputStr  (str) : string for removing all the special character in it.

    Returns :
        outputStr (str) : returns the alphanumeric string.
    """

    if not inputStr:
        inputStr = gui_get_any_input_from_user('input string to remove Special characters')

    if inputStr:
        outputStr = ''.join(e for e in inputStr if e.isalnum())
        return outputStr  

# @background
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
    
    try:
        wb = Workbook()
        ws = wb.active
        ws.title =sheet_name

        if not fullPathToTheFolder:
            fullPathToTheFolder = gui_get_folder_path_from_user('the folder to create excel file')
            
        if not excelFileName:
            excelFileName = gui_get_any_input_from_user("excel file name (without extension)")
        
        folder_create(fullPathToTheFolder)

        if ".xlsx" in excelFileName:
            excel_path = os.path.join(fullPathToTheFolder,excelFileName)
        else:
            excel_path = os.path.join(fullPathToTheFolder,excelFileName + ".xlsx")
            
        excel_path = Path(excel_path)

        wb.save(filename = excel_path)
        
        return True
    except Exception as ex:
        print("Error in excel_create_excel_file_in_given_folder="+str(ex))

def folder_create_text_file(textFolderPath="",txtFileName=""):
    """
    Creates Text file in the given path.
    Internally this uses folder_create() method to create folders if the folder/s does not exist.
    automatically adds txt extension if not given in textFilePath.

    Parameters:
        textFilePath (str) : Complete path to the folder with double slashes.
    """
    try:

        if not textFolderPath:
            textFolderPath = gui_get_folder_path_from_user('the folder to create text file')
        
        if not txtFileName:
            txtFileName = gui_get_any_input_from_user("text file name")
            txtFileName = txtFileName 

        if ".txt" not in txtFileName:
            txtFileName = txtFileName + ".txt"
            
        file_path = os.path.join(textFolderPath, txtFileName)
        file_path = Path(file_path)
        
        f = open(file_path, 'w',encoding="utf-8")
        f.close()
        
    except Exception as ex:
        print("Error in folder_create_text_file="+str(ex))

def excel_if_value_exists(excel_path="",sheet_name='Sheet1',header=0,usecols="",value=""):
    """
    Check if a given value exists in given excel. Returns True / False
    """
    try:
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user('to search the VALUE')

        if not value:
            value = gui_get_any_input_from_user('VALUE to be searched')
        
        if usecols:
            df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header, usecols=usecols)
        else:
            df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header)
        
        if value in df.values:
            df = ''
            return True
        else:
            df = ''
            return False

    except Exception as ex:
        print("Error in excel_if_value_exists="+str(ex))
        
# WatchDog : Monitors the given folder for creation / modification / deletion 
class FileMonitor_Handler(watchdog.events.PatternMatchingEventHandler):
    file_path = ""
    def __init__(self):
        watchdog.events.PatternMatchingEventHandler.__init__(self, ignore_patterns = None,
                                                     ignore_directories = False, case_sensitive = True)
    
    def on_created(self, event):
        file_path = Path(str(event.src_path))

        print("Created : {}".format(file_path))
             
    def on_deleted(self, event):
        file_path = Path(str(event.src_path))
        print("Deleted : {}".format(file_path))

    def on_modified(self,event):
        file_path = Path(str(event.src_path))
        print("Modified : {}".format(file_path))

def create_batch_file(application_exe_pyw_file_path=""):
    """
    Creates .bat file for the given application / exe or even .pyw BOT developed by you. This is required in Task Scheduler.
    """

    global batch_file_path
    try:
        if not application_exe_pyw_file_path:
            
            application_exe_pyw_file_path = gui_get_any_file_from_user('.pyw/.exe file for which .bat is to be made')

            while not (str(application_exe_pyw_file_path).endswith(".exe") or str(application_exe_pyw_file_path).endswith(".pyw")):
                print("Please choose the file ending with .pyw or .exe")
                application_exe_pyw_file_path = gui_get_any_file_from_user('.pyw/.exe file for which .bat is to be made')
            
        application_name= ""

        if str(application_exe_pyw_file_path).endswith(".exe"):
            application_name = _extract_filename_from_filepath(application_exe_pyw_file_path) + ".exe"
        else:
            application_name = _extract_filename_from_filepath(application_exe_pyw_file_path) + ".pyw"

        cmd = ""

        if "exe" in application_name:
            application_name = str(application_name).replace("exe","bat")
            cmd = "start \"\" " + '"' + application_exe_pyw_file_path + '" /popup\n'

        elif "pyw" in application_name: 
            application_name = str(application_name).replace("pyw","bat")
            cmd = "start \"\" " + '"' + sys.executable + '" ' + '"' + application_exe_pyw_file_path + '" /popup\n'

        batch_file_path = os.path.join(batch_file_path,application_name)
        batch_file_path = Path(batch_file_path)
        
        if not os.path.exists(batch_file_path):
            
            f = open(batch_file_path, 'w',encoding="utf-8")
            f.write("@ECHO OFF\n")
            f.write("timeout 5 > nul\n")
            f.write(cmd) 
            f.write("exit")    
            f.close()

        print("Batch file saved in " + str(batch_file_path))
    except Exception as ex:
        print("Error in create_batch_file="+str(ex))
    
def excel_create_file(fullPathToTheFile="",fileName="",sheet_name="Sheet1"):
    try:
        if not fullPathToTheFile:
            fullPathToTheFile = gui_get_any_input_from_user('folder path to create excel')

        if not fileName:
            fileName = gui_get_any_input_from_user("Excel File Name (without extension)")

        if not os.path.exists(fullPathToTheFile):
            os.makedirs(fullPathToTheFile)
        if ".xlsx" not in fileName:
            fileName = fileName + ".xlsx"

        wb = Workbook()
        ws = wb.active
        ws.title =sheet_name

        fileName = os.path.join(fullPathToTheFile,fileName)
        
        fileName = Path(fileName)

        wb.save(filename = fileName)
        
        return True
    except Exception as ex:
        print("Error in excel_create_file="+str(ex))
    
def folder_get_all_filenames_as_list(strFolderPath="",extension='all'):
    """
    Get all the files of the given folder in a list.

    Parameters:
        strFolderPath  (str) : Location of the folder.
        extension      (str) : extention of the file. by default all the files will be listed regarless of the extension.
    
    returns:
        allFilesOfaFolderAsLst (List) : all the file names as a list.
    """
    try:
        if not strFolderPath:
            strFolderPath = gui_get_folder_path_from_user('a folder to get all its filenames')

        if extension == "all":
            allFilesOfaFolderAsLst = [ f for f in os.listdir(strFolderPath)]
        else:
            allFilesOfaFolderAsLst = [ f for f in os.listdir(strFolderPath) if f.endswith(extension) ]

        return allFilesOfaFolderAsLst
    except Exception as ex:
        print("Error in folder_get_all_filenames_as_list="+str(ex))
    
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
    file_extension_with_dot = ''
    try:
        if not fullPathOfTheFolder:
            fullPathOfTheFolder = gui_get_folder_path_from_user('a folder to delete all its files')

        count = 0 
        if "." not in file_extension_without_dot :
            file_extension_with_dot = "." + file_extension_without_dot

        if file_extension_with_dot.lower() == ".all":
            filelist = [ f for f in os.listdir(fullPathOfTheFolder) ]
        else:
            filelist = [ f for f in os.listdir(fullPathOfTheFolder) if f.endswith(file_extension_with_dot) ]
        print(filelist)
        for f in filelist:
            try:
                file_path = os.path.join(fullPathOfTheFolder, f)
                file_path = Path(file_path)
                os.remove(file_path)
                count +=1 
            except:
                pass
        
        return count
    except Exception as ex:
        print("Error in folder_delete_all_files="+str(ex)) 
        return -1
        
def key_hit_enter():
    """
    Enter key will be pressed once.
    """
    time.sleep(0.5)
    kb.press_and_release('enter')
    time.sleep(0.5)

def message_flash(msg="",delay=3):
    """
    specified msg will popup for a specified duration of time with OK button.

    Parameters:
        msg     (str) : message to popup.
        delay   (int) : duration of the popup.
    """
    try:
        if not msg:
            msg = gui_get_any_input_from_user("flash message")

        r = Timer(int(delay), key_hit_enter)
        r.start()
        pg.alert(text=msg, title='ClointFusion', button='OK')
    except Exception as ex:
        print("ERROR in message_flash="+str(ex))

def window_show_desktop():
    """
    Minimizes all the applications and shows Desktop.
    """
    try:
        time.sleep(0.5)
        kb.press_and_release('win+d')
        time.sleep(0.5)
    except Exception as ex:
        print("Error in window_show_desktop="+str(ex))
    
def window_get_all_opened_titles_windows():
    """
    Gives the title of all the existing (open) windows.

    Returns:
        allTitles_lst  (list) : returns all the titles of the window as list.
    """
    try:
        allTitles_lst = []
        lst = gw.getAllTitles()
        for item in lst:
            if str(item).strip() != "" and str(item).strip() not in allTitles_lst:
                allTitles_lst.append(str(item).strip())
        return allTitles_lst
    except Exception as ex:
        print("Error in window_get_all_opened_titles="+str(ex))
    
def _window_find_exact_name(windowName=""):
    """
    Gives you the exact window name you are looking for.

    Parameters:
        windowName  (str) : Name of the window to find.

    Returns:
        win (str)              : Exact window name.
        window_found (boolean) : A boolean TRUE if the window is found
    """
    win = ""
    window_found = False

    if not windowName:
        windowName = gui_get_any_input_from_user("Partial Window Name")

    try:
        lst = gw.getAllTitles()
        
        for item in lst:
            if str(item).strip():
                if str(windowName).lower() in str(item).lower():
                    win = item
                    window_found = True
                    break
        return win, window_found
    except Exception as ex:
        print("Error in _window_find_exact_name="+str(ex))
    
def window_activate_and_maximize_windows(windowName=""):
    """
    Activates and maximizes the desired window.

    Parameters:
        windowName  (str) : Name of the window to maximize.
    """
    try:
        if not windowName:
            open_win_list = window_get_all_opened_titles_windows()
            windowName = gui_get_dropdownlist_values_from_user("window titles to Activate & Maximize",dropdown_list=open_win_list,multi_select=False)[0]
            
        item,window_found = _window_find_exact_name(windowName)
        if window_found:
            windw = gw.getWindowsWithTitle(item)[0]
            windw.activate()
            time.sleep(2)
            windw.maximize()
            time.sleep(2)
        else:
            print("No window OPEN by name="+str(windowName))
    except Exception as ex:
        print("Error in window_activate_and_maximize="+str(ex))
    
def window_minimize_windows(windowName=""):
    """
    Activates and minimizes the desired window.

    Parameters:
        windowName  (str) : Name of the window to miniimize.
    """
    try:
        if not windowName:
            open_win_list = window_get_all_opened_titles_windows()
            windowName = gui_get_dropdownlist_values_from_user("window titles to Minimize",dropdown_list=open_win_list,multi_select=False)[0]
            
        item,window_found = _window_find_exact_name(windowName)
        if window_found:
            windw = gw.getWindowsWithTitle(item)[0]
            windw.minimize()
            time.sleep(1)
        else:
            print("No window available to minimize by name="+str(windowName))
    except Exception as ex:
        print("Error in window_minimize="+str(ex))

def window_close_windows(windowName=""):
    """
    Close the desired window.

    Parameters:
        windowName  (str) : Name of the window to close.
    """
    try:
        if not windowName:
            open_win_list = window_get_all_opened_titles_windows()
            windowName = gui_get_dropdownlist_values_from_user("window titles to Close",dropdown_list=open_win_list,multi_select=False)[0]
            
        item,window_found = _window_find_exact_name(windowName)
        if window_found:
            windw = gw.getWindowsWithTitle(item)[0]
            windw.close()
            time.sleep(1)
        else:
            print("No window available to close, by name="+str(windowName))
    except Exception as ex:
        print("Error in window_close="+str(ex))

def launch_any_exe_bat_application(pathOfExeFile=""):
    """
    Launches any exe or batch file or excel file etc.

    Parameters:
        pathOfExeFile  (str) : location of the file with extension.
    """
    try:
        if not pathOfExeFile:
            pathOfExeFile = gui_get_any_file_from_user('EXE or BAT file')

        try:
            subprocess.Popen(pathOfExeFile)
        except:
            os.startfile(pathOfExeFile)

        time.sleep(2) 
 
        try:
            import win32gui, win32con
            time.sleep(3) 
            hwnd = win32gui.GetForegroundWindow()
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        except Exception as ex1:
            print("launch_any_exe_bat_application"+str(ex1))

        time.sleep(1) 
    except Exception as ex:
        print("ERROR in launch_any_exe_bat_application="+str(ex))

class myThread1 (threading.Thread):
    def __init__(self,err_str):
        threading.Thread.__init__(self)
        self.err_str = err_str

    def run(self):
        message_flash(self.err_str)

class myThread2 (threading.Thread):
    def __init__(self,strFilePath):
        threading.Thread.__init__(self)
        self.strFilePath = strFilePath

    def run(self):
        time.sleep(1)
        img = pg.screenshot()
        time.sleep(1)

        dt_tm= str(datetime.datetime.now())    
    
        dt_tm = dt_tm.replace(" ","_")
        dt_tm = dt_tm.replace(":","-")
        dt_tm = dt_tm.split(".")[0]
        filePath = self.strFilePath + str(dt_tm)  + ".PNG"

        img.save(str(filePath))
        
def take_error_screenshot(err_str):
    """
    Takes screenshot of an error popup parallely without waiting for the flow of the program.
    The screenshot will be saved in the log folder for reference.

    Parameters:
        err_str  (str) : exception.
    """
    global error_screen_shots_path
    try:
        thread1 = myThread1(err_str)
        thread2 = myThread2(error_screen_shots_path)

        thread1.start()
        thread2.start()

        thread1.join()
        thread2.join()
    except Exception as ex:
        print("Error in take_error_screenshot="+str(ex))
    
def update_log_excel_file(message=""):
    """
    Given message will be updated in the excel log file.

    Parameters:
        message  (str) : message to update.

    Retursn:
        returns a boolean true if updated sucessfully
    """
    global status_log_excel_filepath
    try:
        if not message:
            message = gui_get_any_input_from_user("message to Update Log file")

        df = pd.DataFrame({'Timestamp': [datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")], 'Status':[message]})
        writer = pd.ExcelWriter(status_log_excel_filepath, engine='openpyxl')
        writer.book = load_workbook(status_log_excel_filepath)
        writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
    
        reader = pd.read_excel(status_log_excel_filepath)
        df.to_excel(writer,index=False,header=False,startrow=len(reader)+1)
        writer.save()

        return True
    except Exception as ex:
        print("Error in update_log_excel_file="+str(ex))
        return False
    
def string_extract_only_alphabets(inputString=""):
    """
    Returns only alphabets from given input string
    """
    if not inputString:
        inputString = gui_get_any_input_from_user("input string to get only Alphabets")

    outputStr = ''.join(e for e in inputString if e.isalpha())
    return outputStr 

def string_extract_only_numbers(inputString=""):
    """
    Returns only numbers from given input string
    """
    if not inputString:
        inputString = gui_get_any_input_from_user("input string to get only Numbers")

    outputStr = ''.join(e for e in inputString if e.isnumeric())
    return outputStr       
@lru_cache(None)
def call_otsu_threshold(img_title, is_reduce_noise=False):
    """
    OpenCV internal function for OCR
    """
    
    image = cv2.imread(img_title, 0)

    
    if is_reduce_noise:
        image = cv2.GaussianBlur(image, (5, 5), 0)

    
    _ , image_result = cv2.threshold(
        image, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU,
    )
    
    cv2.imwrite(img_title, image_result)
    cv2.destroyAllWindows()

@lru_cache(None)
def read_image_cv2(img_path):
    """
    Saves the image in cv2 format.

    Parameters:
        img_path  (str) : location of the image.
    
    returns:
        image (cv2) : image in cv2 format will be returned.
    """
    if img_path and os.path.exists(img_path):
        try:
            image = cv2.imread(img_path)
            return image
        except Exception as ex:
            print("read_image_cv2 = "+str(ex))
        
    else:
        print("File not found="+str(img_path))


    

    
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
    try:
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user("to get row/column count")
            
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header)
        row, col = df.shape
        row = row + 1
        return row, col
    except Exception as ex:
        print("Error in excel_get_row_column_count="+str(ex))
    
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
    try:
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user('to copy range from')
            
        if startCol == 0 and startRow ==0 and endCol == 0 and endRow == 0:
            sRow_sCol_eRow_Col = gui_get_any_input_from_user('startRow , startCol, endRow, endCol (comma separated, index from 1)')    

            if sRow_sCol_eRow_Col:
                startRow , startCol, endRow, endCol = str(sRow_sCol_eRow_Col).split(",")
                startRow = int(startRow)
                startCol = int(startCol)
                endRow = int(endRow)
                endCol = int(endCol)

        from_wb = load_workbook(filename = excel_path)
        try:
            fromSheet = from_wb[sheet_name]
        except:
            fromSheet = from_wb.worksheets[0]
        rangeSelected = []

        if endRow < startRow:
            endRow = startRow

        #Loops through selected Rows
        for i in range(startRow,endRow + 1,1):
            #Appends the row to a RowSelected list
            rowSelected = []
            for j in range(startCol,endCol+1,1):
                rowSelected.append(fromSheet.cell(row = i, column = j).value)
            #Adds the RowSelected List and nests inside the rangeSelected
            rangeSelected.append(rowSelected)
    
        return rangeSelected
    except Exception as ex:
        print("Error in copy_range_from_excel_sheet="+str(ex))
    
def excel_copy_paste_range_from_to_sheet(excel_path="", sheet_name='Sheet1', startCol=0, startRow=0, endCol=0, endRow=0, copiedData=""):#*
    """
    Pastes the copied data in specific range of the given excel sheet.
    """
    try:
        try:
            if not copiedData:
                copiedData = excel_copy_range_from_sheet()

            if not excel_path:
                excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user('to paste range into')
                
            if startCol == 0 and startRow ==0 and endCol == 0 and endRow == 0:
                sRow_sCol_eRow_Col = gui_get_any_input_from_user('startRow , startCol, endRow, endCol (comma separated, index from 1)')    

                if sRow_sCol_eRow_Col:
                    startRow , startCol, endRow, endCol = str(sRow_sCol_eRow_Col).split(",")
                    startRow = int(startRow)
                    startCol = int(startCol)
                    endRow = int(endRow)
                    endCol = int(endCol)

            to_wb = load_workbook(filename = excel_path)
            toSheet = to_wb[sheet_name]

        except:
            try:
                excel_create_excel_file_in_given_folder((str(excel_path[:(str(excel_path).rindex("\\"))])),(str(excel_path[str(excel_path).rindex("\\")+1:excel_path.find(".")])),sheet_name)
            except:
                excel_create_excel_file_in_given_folder((str(excel_path[:(str(excel_path).rindex("/"))])),(str(excel_path[str(excel_path).rindex("/")+1:excel_path.find(".")])),sheet_name)

            to_wb = load_workbook(filename = excel_path)

            toSheet = to_wb[sheet_name]

        if endRow < startRow:
            endRow = startRow

        countRow = 0
        for i in range(startRow,endRow+1,1):
            countCol = 0
            for j in range(startCol,endCol+1,1):
                toSheet.cell(row = i, column = j).value = copiedData[countRow][countCol]
                countCol += 1
            countRow += 1
        to_wb.save(excel_path)
        return countRow-1
    except Exception as ex:
        print("Error in excel_copy_paste_range_from_to_sheet="+str(ex))
    
def _excel_copy_range(startCol=1, startRow=1, endCol=1, endRow=1, sheet='Sheet1'):
    """
    Copies the specific range from the given excel sheet.
    """
    try:
        rangeSelected = []
        #Loops through selected Rows
        for k in range(startRow,endRow + 1,1):
            #Appends the row to a RowSelected list
            rowSelected = []
            for l in range(startCol,endCol+1,1):
                rowSelected.append(sheet.cell(row = k, column = l).value)
            #Adds the RowSelected List and nests inside the rangeSelected
            rangeSelected.append(rowSelected)

        return rangeSelected

    except Exception as ex:
        print("Error in _excel_copy_range="+str(ex))
    
def _excel_paste_range(startCol=1, startRow=1, endCol=1, endRow=1, sheetReceiving='Sheet1',copiedData=[]):
    """
    Pastes the specific range to the given excel sheet.
    """
    try:
        countRow = 0
        for k in range(startRow,endRow+1,1):
            countCol = 0
            for l in range(startCol,endCol+1,1):
                sheetReceiving.cell(row = k, column = l).value = copiedData[countRow][countCol]
                countCol += 1
            countRow += 1
        return countRow

    except Exception as ex:
        print("Error in _excel_paste_range="+str(ex))

def excel_split_by_column(excel_path="",sheet_name='Sheet1',header=0,columnName=""):#*
    """
    Splits the excel file by Column Name
    """
    try:
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user('to split by column')
            
        if not columnName:
            col_lst = excel_get_all_header_columns(excel_path, sheet_name, header)
            columnName = gui_get_dropdownlist_values_from_user('this list of Columns (to split)',col_lst)

        data_df = pd.read_excel(excel_path,sheet_name=sheet_name,header=header,dtype=str)
        
        grouped_df = data_df.groupby(columnName)
        
        for data in  grouped_df:  
            file_path = os.path.join(output_folder_path,str(data[0]) + ".xlsx")
            file_path = Path(file_path)
            grouped_df.get_group(data[0]).to_excel(file_path, index=False)
            
    except Exception as ex:
        print("Error in excel_split_by_column="+str(ex))
    
def excel_split_the_file_on_row_count(excel_path="", sheet_name = 'Sheet1', rowSplitLimit="", outputFolderPath="", outputTemplateFileName ="Split"):#*
    """
    Splits the excel file as per given row limit
    """
    try:
        if not excel_path:
            excel_path, sheet_name, _ = gui_get_excel_sheet_header_from_user('to split on row count')
            
        if not rowSplitLimit:
            rowSplitLimit = int(gui_get_any_input_from_user("row split Count/Limit Ex: 20"))

        if not outputFolderPath:
            outputFolderPath = gui_get_folder_path_from_user('output folder to Save split excel files')

        src_wb = op.load_workbook(excel_path)
        src_ws = src_wb.worksheets[0] 

        src_ws_max_rows = src_ws.max_row
        src_ws_max_cols= src_ws.max_column 

        i = 1
        start_row = 2

        while start_row <= src_ws_max_rows:
            
            dest_wb = Workbook()
            dest_ws = dest_wb.active
            dest_ws.title = sheet_name

            #Copy ROW-1 (Header) from SOURCE to Each DESTINATION file
            selectedRange = _excel_copy_range(1,1,src_ws_max_cols,1,src_ws) #startCol, startRow, endCol, endRow, sheet
            _ =_excel_paste_range(1,1,src_ws_max_cols,1,dest_ws,selectedRange) #startCol, startRow, endCol, endRow, sheetReceiving,copiedData
            
            selectedRange = ""
            selectedRange = _excel_copy_range(1,start_row,src_ws_max_cols,start_row + rowSplitLimit - 1,src_ws) #startCol, startRow, endCol, endRow, sheet   
            _ =_excel_paste_range(1,2,src_ws_max_cols,rowSplitLimit + 1,dest_ws,selectedRange) #startCol, startRow, endCol, endRow, sheetReceiving,copiedData

            start_row = start_row + rowSplitLimit

            try:
                dest_file_name = str(outputFolderPath) + "\\" + outputTemplateFileName + "-" + str(i) + ".xlsx"
            except:
                dest_file_name = str(outputFolderPath) + "/" + outputTemplateFileName + "-" + str(i) + ".xlsx"

            dest_file_name = Path(dest_file_name)
            dest_wb.save(dest_file_name)
            
            i = i + 1
        return True
    except Exception as ex:
        print("Error in excel_split_the_file_on_row_count="+str(ex))
    
def excel_merge_all_files(input_folder_path="",output_folder_path=""):
    """
    Merges all the excel files in the given folder
    """
    try:
        if not input_folder_path:
            input_folder_path = gui_get_folder_path_from_user('input folder to MERGE files from')

        if not output_folder_path:
            output_folder_path = gui_get_folder_path_from_user('output folder to store Final merged file')
        
        filelist = [ f for f in os.listdir(input_folder_path) if f.endswith(".xlsx") ]
        all_excel_file_lst = []
        for file1 in filelist:
            file_path = os.path.join(input_folder_path,file1)
            file_path = Path(file_path)
            
            all_excel_file = pd.read_excel(file_path,dtype=str)
            all_excel_file_lst.append(all_excel_file)

        appended_df = pd.concat(all_excel_file_lst)
        time_stamp_now=datetime.datetime.now().strftime("%m-%d-%Y")
        final_path = os.path.join(output_folder_path, "Final-" + time_stamp_now + ".xlsx")
        final_path= Path(final_path)
        appended_df.to_excel(final_path, index=False)
        
        return True
    except Exception as ex:
        print("Error in excel_merge_all_files="+str(ex))

def excel_drop_columns(excel_path="", sheet_name='Sheet1', header=0, columnsToBeDropped = ""):
    """
    Drops the desired column from the given excel file
    """
    try:
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user('input excel to Drop the columns from')

        if not columnsToBeDropped:
            col_lst = excel_get_all_header_columns(excel_path, sheet_name, header)
            columnsToBeDropped = gui_get_dropdownlist_values_from_user('columns list to drop',col_lst) 

        df=pd.read_excel(excel_path,sheet_name=sheet_name, header=header) 

        if isinstance(columnsToBeDropped, list):
            df.drop(columnsToBeDropped, axis = 1, inplace = True) 
        else:
            df.drop([columnsToBeDropped], axis = 1, inplace = True) 


            
        with pd.ExcelWriter(excel_path) as writer:
            df.to_excel(writer, sheet_name=sheet_name,index=False) 

    except Exception as ex:
        print("Error in excel_drop_columns="+str(ex))

def excel_sort_columns(excel_path="",sheet_name='Sheet1',header=0,firstColumnToBeSorted=None,secondColumnToBeSorted=None,thirdColumnToBeSorted=None,firstColumnSortType=True,secondColumnSortType=True,thirdColumnSortType=True):#*
    """
    A function which takes excel full path to excel and column names on which sort is to be performed

    """
    try:
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user('to sort the column')

        if not firstColumnToBeSorted:
            col_lst = excel_get_all_header_columns(excel_path, sheet_name, header)
            usecols = gui_get_dropdownlist_values_from_user('minimum 1 and maximum 3 columns to sort',col_lst)
            
            if len(usecols) == 3:
                firstColumnToBeSorted , secondColumnToBeSorted , thirdColumnToBeSorted = usecols
            elif len(usecols) == 2:
                firstColumnToBeSorted , secondColumnToBeSorted = usecols
            elif len(usecols) == 1:
                firstColumnToBeSorted = usecols[0]
        df=pd.read_excel(excel_path,sheet_name=sheet_name, header=header)
        if thirdColumnToBeSorted is not None and secondColumnToBeSorted is not None and firstColumnToBeSorted is not None:
            df=df.sort_values([firstColumnToBeSorted,secondColumnToBeSorted,thirdColumnToBeSorted],ascending=[firstColumnSortType,secondColumnSortType,thirdColumnSortType])
        
        elif secondColumnToBeSorted is not None and firstColumnToBeSorted is not None:
            df=df.sort_values([firstColumnToBeSorted,secondColumnToBeSorted],ascending=[firstColumnSortType,secondColumnSortType])
        
        elif firstColumnToBeSorted is not None:
            df=df.sort_values([firstColumnToBeSorted],ascending=[firstColumnSortType])

        writer = pd.ExcelWriter(excel_path, engine='openpyxl')
        writer.book = load_workbook(excel_path)
        writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
    
        df.to_excel(writer,sheet_name=sheet_name,index=False)

        writer.save()
        
        return True
    except Exception as ex:
        print("Error in excel_sort_columns="+str(ex))        
    
def excel_clear_sheet(excel_path="",sheet_name="Sheet1", header=0):
    """
    Clears the contents of given excel files keeping header row intact
    """
    try:
        
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user('to clear the sheet')

        df = pd.read_excel(excel_path,sheet_name=sheet_name,header=header) 
        df = df.head(0)

        with pd.ExcelWriter(excel_path) as writer:
            df.to_excel(writer,sheet_name=sheet_name, index=False)

    except Exception as ex:
        print("Error in excel_clear_sheet="+str(ex))

def excel_set_single_cell(excel_path="", sheet_name="Sheet1", header=0, columnName="", cellNumber=0, setText=""): #*
    """
    Writes the given text to the desired column/cell number for the given excel file
    """
    try:
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user('to set cell')
            
        if not columnName:
            col_lst = excel_get_all_header_columns(excel_path, sheet_name, header)
            columnName = gui_get_dropdownlist_values_from_user('list of columns to set vlaue',col_lst,multi_select=False)   

        if not setText:
            setText = gui_get_any_input_from_user("text value to set the cell")

        df = pd.read_excel(excel_path,sheet_name=sheet_name,header=header)
        
        writer = pd.ExcelWriter(excel_path, engine='openpyxl')
        writer.book = load_workbook(excel_path)
        writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
        
        df.at[cellNumber,columnName] = setText
        df.to_excel(writer, sheet_name=sheet_name ,index=False)    
        writer.save()
        
        return True

    except Exception as ex:
        print("Error in excel_set_single_cell="+str(ex))

def excel_get_single_cell(excel_path="",sheet_name="Sheet1",header=0, columnName="",cellNumber=0): #*
    """
    Gets the text from the desired column/cell number of the given excel file
    """
    try:
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user('to get cell')
            
        if not columnName:
            col_lst = excel_get_all_header_columns(excel_path, sheet_name, header)
            columnName = gui_get_dropdownlist_values_from_user('list of columns to get vlaue',col_lst,multi_select=False)   

        if not isinstance(columnName, list):
            columnName = [columnName]       

        df = pd.read_excel(excel_path,sheet_name=sheet_name,header=header,usecols={columnName[0]})
        cellValue = df.at[cellNumber,columnName[0]]
        return cellValue
    except Exception as ex:
        print("Error in excel_get_single_cell="+str(ex))

def excel_remove_duplicates(excel_path="",sheet_name="Sheet1", header=0, columnName="", saveResultsInSameExcel=True, which_one_to_keep="first"): #*
    """
    Drops the duplicates from the desired Column of the given excel file
    """
    try:
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user('to remove duplicates')
            
        if not columnName:
            col_lst = excel_get_all_header_columns(excel_path, sheet_name, header)
            columnName = gui_get_dropdownlist_values_from_user('list of columns to remove duplicates',col_lst)  
    
        df = pd.read_excel(excel_path, sheet_name=sheet_name,header=header) 

        count = 0 
        if saveResultsInSameExcel:
            df.drop_duplicates(subset=columnName, keep=which_one_to_keep, inplace=True)
            with pd.ExcelWriter(excel_path) as writer:
                df.to_excel(writer,sheet_name=sheet_name,index=False)

            count = df.shape[0]
        else:
            df1 = df.drop_duplicates(subset=columnName, keep=which_one_to_keep, inplace=False)
            excel_path = str(excel_path).replace(".","_DupDropped.")
            with pd.ExcelWriter(excel_path) as writer:
                df1.to_excel(writer,sheet_name=sheet_name,index=False)
            count = df1.shape[0]

        return count
    except Exception as ex:
        print("Error in excel_remove_duplicates="+str(ex))
    
def excel_vlook_up(filepath_1="", sheet_name_1 = 'Sheet1', header_1 = 0, filepath_2="", sheet_name_2 = 'Sheet1', header_2 = 0, Output_path="", OutputExcelFileName="", match_column_name="",how='left'):#*
    """
    Performs excel_vlook_up on the given excel files for the desired columns. Possible values for how are "inner","left", "right", "outer"
    """
    try:
        if not filepath_1:
            filepath_1, sheet_name_1, header_1 = gui_get_excel_sheet_header_from_user('(Vlookup) first excel')
             
        if not filepath_2:
            filepath_2, sheet_name_2, header_2 = gui_get_excel_sheet_header_from_user('(Vlookup) second excel')
            
        if not match_column_name:
            col_lst = excel_get_all_header_columns(filepath_1, sheet_name_1, header_1)
            match_column_name = gui_get_dropdownlist_values_from_user('Vlookup column name to be matched',col_lst,multi_select=False) 
            match_column_name = match_column_name[0]
        df1 = pd.read_excel(filepath_1, sheet_name = sheet_name_1, header = header_1)
        df2 = pd.read_excel(filepath_2, sheet_name = sheet_name_2, header = header_2)

        df = pd.merge(df1, df2, on= match_column_name, how = how)

        output_file_path = ""
        if str(OutputExcelFileName).endswith(".*"):
            OutputExcelFileName = OutputExcelFileName.split(".")[0]
        
        if Output_path and OutputExcelFileName:
            if ".xlsx" in OutputExcelFileName:
                output_file_path = os.path.join(Output_path, OutputExcelFileName)
            else:
                output_file_path = os.path.join(Output_path, OutputExcelFileName  + ".xlsx")

        else:
            output_file_path = filepath_1

        output_file_path = Path(output_file_path)
        with pd.ExcelWriter(output_file_path) as writer:
            df.to_excel(writer, index=False)

        return True
    
    except Exception as ex:
        print("Error in excel_vlook_up="+str(ex))
    
def screen_clear_search(delay=0.2):
    """
    Clears previously found text (crtl+f highlight)
    """
    try:
        kb.press_and_release("ctrl+f")
        time.sleep(delay)
        pg.typewrite("^%#")
        time.sleep(delay)
        kb.press_and_release("esc")
        time.sleep(delay)
    except Exception as ex:
        print("Error in screen_clear_search="+str(ex))
    
def scrape_save_contents_to_notepad(folderPathToSaveTheNotepad="",X=pg.size()[0]/2,Y=pg.size()[1]/2): #"Full path to the folder (with double slashes) where notepad is to be stored"
    """
    Copy pastes all the available text on the screen to notepad and saves it.
    """
    try:
        if not folderPathToSaveTheNotepad:
            folderPathToSaveTheNotepad = gui_get_folder_path_from_user('folder to save notepad contents')

        message_counter_down_timer(3)
        time.sleep(1)
        
        pg.click(X,Y)
        
        time.sleep(0.5)

        kb.press_and_release("ctrl+a")
        time.sleep(1)
        kb.press_and_release("ctrl+c")
        time.sleep(1)
        
        clipboard_data = clipboard.paste()
        time.sleep(2)
        
        screen_clear_search()

        notepad_file_path = Path(folderPathToSaveTheNotepad)    
        notepad_file_path = notepad_file_path / 'notepad-contents.txt'

        f = open(notepad_file_path, "w", encoding="utf-8")
        f.write(clipboard_data)
        time.sleep(10)
        f.close()

        clipboard_data = ''
        return "Saved the contents at " + str(notepad_file_path)
    except Exception as ex:
        print("Error in scrape_save_contents_to_notepad = "+str(ex))
    
def scrape_get_contents_by_search_copy_paste(highlightText=""):
    """
    Gets the focus on the screen by searching given text using crtl+f and performs copy/paste of all data. Useful in Citrix applications
    This is useful in Citrix applications
    """
    output_lst_newline_removed = []
    try:
        if not highlightText:
            highlightText = gui_get_any_input_from_user("text to be searched in Citrix environment")

        time.sleep(1)
        kb.press_and_release("ctrl+f")
        time.sleep(1)
        pg.typewrite(highlightText)
        time.sleep(1)
        kb.press_and_release("enter")
        time.sleep(1)
        kb.press_and_release("esc")
        time.sleep(2)

        pg.PAUSE = 2
        kb.press_and_release("ctrl+a")
        time.sleep(2)
        kb.press_and_release("ctrl+c")
        time.sleep(2)
        
        clipboard_data = clipboard.paste()
        time.sleep(2)
        
        screen_clear_search()

        entire_data_as_list= clipboard_data.splitlines()
        for line in entire_data_as_list:
            if line.strip():
                output_lst_newline_removed.append(line.strip())

        clipboard_data = ''
        return output_lst_newline_removed
    except Exception as ex:
        print("Error in scrape_get_contents_by_search_copy_paste="+str(ex))
    
def mouse_move(x="",y=""):
    """
    Moves the cursor to the given X Y Co-ordinates.
    """
    try:
        if not x and not y:
            x_y = str(gui_get_any_input_from_user("X,Y co-ordinates to the move Mouse to. Ex: 200,215"))
            if "," in x_y:
                x, y = x_y.split(",")
                x = int(x)
                y = int(y)
            else:
                x = x_y.split(" ")[0]
                y = x_y.split(" ")[1]
        if x and y:
            time.sleep(0.2)
            pg.moveTo(x,y)
            time.sleep(0.2)
    except Exception as ex:
        print("Error in mouse_move="+str(ex))
    
def mouse_get_color_by_position(pos=[]):
    """
    Gets the color by X Y co-ordinates of the screen.
    """
    try:
        if not pos:
            pos1 = gui_get_any_input_from_user("X,Y co-ordinates to get its color. Ex: 200,215")
            pos = tuple(map(int, pos1.split(',')))

        im = pg.screenshot()
        time.sleep(0.5)
        return im.getpixel(pos)    
    except Exception as ex:
        print("Error in mouse_get_color_by_position = "+str(ex))
    
def mouse_click(x="", y="", left_or_right="left", single_double_triple="single", copyToClipBoard_Yes_No="no"):
    """
    Clicks at the given X Y Co-ordinates on the screen using ingle / double / tripple click(s).
    Optionally copies selected data to clipboard (works for double / triple clicks)
    """
    try:
        if not x and not y:
            x_y = str(gui_get_any_input_from_user("X,Y co-ordinates to perform Mouse (Left) Click. Ex: 200,215"))
            if "," in x_y:
                x, y = x_y.split(",")
                x = int(x)
                y = int(y)
            else:
                x = int(x_y.split(" ")[0])
                y = int(x_y.split(" ")[1])

        copiedText = ""
        time.sleep(1)

        if x and y:
            if single_double_triple.lower() == "single" and left_or_right.lower() == "left":
                pg.click(x,y)
            elif single_double_triple.lower() == "double" and left_or_right.lower() == "left":
                pg.doubleClick(x,y)
            elif single_double_triple.lower() == "triple" and left_or_right.lower() == "left":
                pg.tripleClick(x,y)
            elif single_double_triple.lower() == "single" and left_or_right.lower() == "right":
                pg.rightClick(x,y)
            time.sleep(1)    

            if copyToClipBoard_Yes_No.lower() == "yes":
                kb.press_and_release("ctrl+c")
                time.sleep(1)
                copiedText = clipboard.paste().strip()
                time.sleep(1)
                
            time.sleep(1)    
            return copiedText
    except Exception as ex:
        print("Error in mouseClick="+str(ex))
    
def mouse_drag_from_to(X1="",Y1="",X2="",Y2="",delay=0.5):
    """
    Clicks and drags from X1 Y1 co-ordinates to X2 Y2 Co-ordinates on the screen
    """
    try:
        if not X1 and not Y1:
            x_y = str(gui_get_any_input_from_user("Mouse Drag FROM Values ex: 200,215"))
            if "," in x_y:
                X1, Y1 = x_y.split(",")
                X1 = int(X1)
                Y1 = int(Y1)

        if not X2 and not Y2:
            x_y = str(gui_get_any_input_from_user("Mouse Drag TO Values ex: 200,215"))
            if "," in x_y:
                X2, Y2 = x_y.split(",")
                X2 = int(X2)
                Y2 = int(Y2)
        time.sleep(0.2)
        pg.moveTo(X1,Y1,duration=delay)
        pg.dragTo(X2,Y2,duration=delay,button='left')
        time.sleep(0.2)
    except Exception as ex:
        print("Error in mouse_drag_from_to="+str(ex))
    
def search_highlight_tab_enter_open(searchText="",hitEnterKey="Yes",shift_tab='No'):
    """
    Searches for a text on screen using crtl+f and hits enter.
    This function is useful in Citrix environment
    """
    try:
        if not searchText:
            searchText = gui_get_any_input_from_user("Search Text to Highlight (in Citrix Environment)")

        time.sleep(0.5)
        kb.press_and_release("ctrl+f")
        time.sleep(0.5)
        kb.write(searchText)
        time.sleep(0.5)
        kb.press_and_release("enter")
        time.sleep(0.5)
        kb.press_and_release("esc")
        time.sleep(0.2)
        if hitEnterKey.lower() == "yes" and shift_tab.lower() == "yes":
            kb.press_and_release("tab")
            time.sleep(0.3)
            kb.press_and_release("shift+tab")
            time.sleep(0.3)
            kb.press_and_release("enter")
            time.sleep(2)
        elif hitEnterKey.lower() == "yes" and shift_tab.lower() == "no":
            kb.press_and_release("enter")
            time.sleep(2)
        return True

    except Exception as ex:
        print("Error in search_highlight_tab_enter_open="+str(ex))
    
def key_press(strKeys=""):
    """
    Emulates the given keystrokes.
    """
    try:
        if not strKeys:            
            strKeys = gui_get_any_input_from_user("keys combination using + as delimeter. Ex: ctrl+O")

        strKeys = strKeys.lower()
        if "shift" in strKeys:
            strKeys = strKeys.replace("shift","left shift+right shift")

        time.sleep(0.5)
        kb.press_and_release(strKeys)
        time.sleep(0.5)
    except Exception as ex:
        print("Error in key_press="+str(ex))
    
def key_write_enter(strMsg="",delay=1,key="e"):
    """
    Writes/Types the given text and press enter (by default) or tab key.
    """
    try:
        if not strMsg:
            strMsg = gui_get_any_input_from_user("message / username / any text")

        time.sleep(0.2)
        kb.write(strMsg)
        time.sleep(delay)
        if key.lower() == "e":
            key_press('enter')
        elif key.lower() == "t":
            key_press('tab')
        time.sleep(1)
    except Exception as ex:
        print("Error in key_write_enter="+str(ex))

def date_convert_to_US_format(input_str=""):
    """
    Converts the given date to US date format.
    """
    try:
        if not input_str:
            input_str = gui_get_any_input_from_user('Date value Ex: 01/01/2021')
        match = re.search(r'\d{4}-\d{2}-\d{2}', input_str) #1
        if match == None:
            match = re.search(r'\d{2}-\d{2}-\d{4}', input_str) #2
            if match == None:
                match = re.search(r'\d{2}/\d{2}/\d{4}', input_str) #3
                if match == None:
                    match = re.search(r'\d{4}/\d{2}/\d{2}', input_str) #4
                    if match == None:
                        match = re.findall(r'(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s\d\d,\s\d{4}',input_str) #5
                        dt=datetime.datetime.strptime(match[0], '%b %d, %Y').date() #5 Jan 01, 2020
                    else:    
                        dt=datetime.datetime.strptime(match.group(), '%Y/%m/%d').date() #4
                else:
                    try:
                        dt=datetime.datetime.strptime(match.group(),'%d/%m/%Y').date() #3
                    except:
                        dt=datetime.datetime.strptime(match.group(),'%m/%d/%Y').date() #3
            else:
                try:
                    dt=datetime.datetime.strptime(match.group(), '%d-%m-%Y').date()#2
                except:
                    dt=datetime.datetime.strptime(match.group(), '%m-%d-%Y').date()#2
        else:
            dt=datetime.datetime.strptime(match.group(), '%Y-%m-%d').date() #1
        return dt.strftime('%m/%d/%Y')    
    except Exception as ex:
        print("Error in date_convert_to_US_format="+str(ex))
    
def mouse_search_snip_return_coordinates_x_y(img="", conf=0.9, wait=180,region=(0,0,pg.size()[0],pg.size()[1])):
    """
    Searches the given image on the screen and returns its center of X Y co-ordinates.
    """ 
    try:
        if not img:
            img = gui_get_any_file_from_user("snip image file, to get X,Y coordinates","png")

        time.sleep(1)

        pos = pg.locateOnScreen(img,confidence=conf,region=region) 
        i = 0
        while pos == None and i < int(wait):
            pos = ()
            pos = pg.locateOnScreen(img, confidence=conf,region=region)   
            time.sleep(1)
            i = i + 1

        time.sleep(1)

        if pos:
            x,y = pos.left + int(pos.width / 2), pos.top + int(pos.height / 2)
            pos = ()
            pos=(x,y)
            
            return pos
        return pos
    except Exception as ex:
        print("Error in mouse_search_snip_return_coordinates_x_y="+str(ex))

def mouse_search_snips_return_coordinates_x_y(img_lst=[], conf=0.9, wait=180, region=(0,0,pg.size()[0],pg.size()[1])):
    """
    Searches the given set of images on the screen and returns its center of X Y co-ordinates of FIRST OCCURANCE
    """ 
    try:
        if not img_lst:
            img_lst_folder_path = gui_get_folder_path_from_user("folder having snip image files, to get X,Y coordinates of any one")

            for img_file in img_lst:
                img_file = os.path.join(img_lst_folder_path,img_file)

                img_file = Path(str(img_file))

                img_lst.append(img_file)

        time.sleep(1)

        if len(img_lst) > 0:
            #Logic = Locate Image Immediately
            pos = ()
            for img in img_lst:
                pos = pg.locateOnScreen(img,confidence=conf,region=region) 
                if pos != None:
                    break

            #Logic = Locate Image with Delay
            i = 0
            while pos == None and i < int(wait):
                pos = ()

                for img in img_lst:
                    pos = pg.locateOnScreen(img,confidence=conf,region=region) 
                    if pos != None:
                        break

                time.sleep(1)
                i = i + 1

            time.sleep(1)

            if pos:
                x,y = pos.left + int(pos.width / 2), pos.top + int(pos.height / 2)
                pos = ()
                pos=(x,y)
                
                return pos
            return pos
        
    except Exception as ex:
        print("Error in mouse_search_snips_return_coordinates_x_y="+str(ex))

def find_text_on_screen(searchText="",delay=0.1, occurance=1,isSearchToBeCleared=False):
    """
    Clears previous search and finds the provided text on screen.
    """
    screen_clear_search() #default

    if not searchText:
        searchText = gui_get_any_input_from_user("search text to Find on screen")

    time.sleep(delay)
    kb.press_and_release("ctrl+f")
    time.sleep(delay)
    pg.typewrite(searchText)
    time.sleep(delay)

    for i in range(occurance-1):
        kb.press_and_release("enter")
        time.sleep(delay)

    kb.press_and_release("esc")
    time.sleep(delay)

    if isSearchToBeCleared:
        screen_clear_search()

def mouse_search_snip_return_coordinates_box(img="", conf=0.9, wait=180,region=(0,0,pg.size()[0],pg.size()[1])):
    """
    Searches the given image on the screen and returns the 4 bounds co-ordinates (x,y,w,h)
    """
    try:
        if not img:
            img = gui_get_any_file_from_user("snip image file, to get BOX coordinates","png")
        time.sleep(1)
        
        pos = pg.locateOnScreen(img,confidence=conf,region=region) 
        i = 0
        while pos == None and i < int(wait):
            pos = ()
            pos = pg.locateOnScreen(img, confidence=conf,region=region)   
            time.sleep(1)
            i = i + 1
        time.sleep(1)
        return pos

    except Exception as ex:
        print("Error in mouse_search_snip_return_coordinates_box="+str(ex))

def mouse_find_highlight_click(searchText="",delay=0.1,occurance=1,left_right="left",single_double_triple="single",copyToClipBoard_Yes_No="no"):
    """
    Searches the given text on the screen, highlights and clicks it.
    """  
    try:
        if not searchText:
            searchText = gui_get_any_input_from_user("search text to Highlight & Click")


        time.sleep(0.2)

        find_text_on_screen(searchText,delay=delay,occurance=occurance,isSearchToBeCleared = True) #clear the search

        img = pg.screenshot()
        img.save(ss_path_b)
        time.sleep(0.2)
        imageA = cv2.imread(ss_path_b)
        time.sleep(0.2)

        find_text_on_screen(searchText,delay=delay,occurance=occurance,isSearchToBeCleared = False) #dont clear the searched text

        img = pg.screenshot()
        img.save(ss_path_a)
        time.sleep(0.2)
        imageB = cv2.imread(ss_path_a)
        time.sleep(0.2)

        # convert both images to grayscale
        grayA = cv2.cvtColor(imageA, cv2.COLOR_BGR2GRAY)
        grayB = cv2.cvtColor(imageB, cv2.COLOR_BGR2GRAY)

        # compute the Structural Similarity Index (SSIM) between the two
        (_, diff) = structural_similarity(grayA, grayB, full=True)
        diff = (diff * 255).astype("uint8")

        thresh = cv2.threshold(diff, 0, 255,
            cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]
        cnts = cv2.findContours(thresh.copy(), cv2.RETR_EXTERNAL,
            cv2.CHAIN_APPROX_SIMPLE)
        cnts = imutils.grab_contours(cnts)

        # loop over the contours
        for c in cnts:
            (x, y, w, h) = cv2.boundingRect(c)
            
            X = int(x + (w/2))
            Y = int(y + (h/2))
            
            mouse_click(x=X,y=Y,left_or_right=left_right,single_double_triple=single_double_triple,copyToClipBoard_Yes_No=copyToClipBoard_Yes_No)
            time.sleep(0.5)
            break

    except Exception as ex:
        print("Error in mouse_find_highlight_click="+str(ex))
                
def schedule_create_task_windows(Weekly_Daily="D",week_day="Sun",start_time_hh_mm_24_hr_frmt="11:00"):#*
    """
    Schedules (weekly & daily options as of now) the current BOT (.bat) using Windows Task Scheduler. Please call create_batch_file() function before using this function to convert .pyw file to .bat
    """
    global batch_file_path
    try:

        str_cmd = ""

        if not batch_file_path:
            batch_file_path = gui_get_any_file_from_user('BATCH file to Schedule. Please call create_batch_file() to create one')

        if Weekly_Daily == "D":
            str_cmd = r"powershell.exe Start-Process schtasks '/create  /SC DAILY /tn ClointFusion\{} /tr {} /st {}' ".format(bot_name,batch_file_path,start_time_hh_mm_24_hr_frmt)
        elif Weekly_Daily == "W":
            str_cmd = r"powershell.exe Start-Process schtasks '/create  /SC WEEKLY /D {} /tn ClointFusion\{} /tr {} /st {}' ".format(week_day,bot_name,batch_file_path,start_time_hh_mm_24_hr_frmt)

        subprocess.call(str_cmd)
        print("Task Scheduled")
    except Exception as ex:
        print("Error in schedule_create_task_windows="+str(ex))

def schedule_delete_task_windows():
    """
    Deletes already scheduled task. Asks user to supply task_name used during scheduling the task. You can also perform this action from Windows Task Scheduler.
    """
    try:
        str_cmd = r"powershell.exe Start-Process schtasks '/delete /tn ClointFusion\{} ' ".format(bot_name)
        
        subprocess.call(str_cmd)
        print("Task {} Deleted".format(bot_name))

    except Exception as ex:
        print("Error in schedule_delete_task="+str(ex))

@lru_cache(None)
def _get_tabular_data_from_website(Website_URL):
    """
    internal function
    """
    all_tables = ""
    try:
        all_tables = pd.read_html(Website_URL)
        return all_tables
    except Exception as ex:
        print("Error in _get_tabular_data_from_website="+str(ex))
    finally:
        return all_tables

def browser_get_html_tabular_data_from_website(Website_URL="",table_index=-1,drop_first_row=False,drop_first_few_rows=[0],drop_last_row=False):
    """
    Web Scrape HTML Tables : Gets Website Table Data Easily as an Excel using Pandas. Just pass the URL of Website having HTML Tables.
    If there are 5 tables on that HTML page and you want 4th table, pass table_index as 3

    Ex: browser_get_html_tabular_data_from_website(Website_URL=URL)
    """
    try:
        if not Website_URL:            
            Website_URL= gui_get_any_input_from_user("website URL to get HTML Tabular Data ex: https://www.google.com ")

        all_tables = _get_tabular_data_from_website(Website_URL)

        if all_tables:
            
            # if no table_index is specified, then get all tables in output
            if table_index == -1:
                try:
                    strFileName = Website_URL[Website_URL.rindex("\\")+1:] + "_All_Tables" +  ".xlsx"
                except:
                    strFileName = Website_URL[Website_URL.rindex("/")+1:] + "_All_Tables" +  ".xlsx"

                excel_create_excel_file_in_given_folder(output_folder_path,strFileName)
            else:
                try:
                    strFileName = Website_URL[Website_URL.rindex("\\")+1:] + "_" + str(table_index) +  ".xlsx"
                except:
                    strFileName = Website_URL[Website_URL.rindex("/")+1:] + "_" + str(table_index) +  ".xlsx"

            strFileName = os.path.join(output_folder_path,strFileName)
            strFileName = Path(strFileName)
            
            if table_index == -1:
                for i in range(len(all_tables)):
                    table = all_tables[i] #lool thru table_index values

                    table = table.reset_index(drop=True) #Avoid multi index error in our dataframes

                    with pd.ExcelWriter(strFileName) as writer:
                        table.to_excel(writer, sheet_name=str(i)) #index=False
            else:
                table = all_tables[table_index] #get required table_index
                
                if drop_first_row:
                    table = table.drop(drop_first_few_rows) # Drop first few rows (passed as list)

                if drop_last_row:
                    table = table.drop(len(table)-1) # Drop last row

            # table.columns = list(table.iloc[0])
            # table = table.drop(len(drop_first_few_rows)) 

                table = table.reset_index(drop=True) 

                table.to_excel(strFileName, index=False)

            print("Table saved as Excel at {} ".format(strFileName))

        else:
            print("No tables found in given website " + str(Website_URL))

    except Exception as ex:
        print("Error in browser_get_html_tabular_data_from_website="+str(ex))



def excel_draw_charts(excel_path="",sheet_name='Sheet1', header=0, x_col="", y_col="", color="", chart_type='bar', title='ClointFusion', show_chart=False):

    """
    Interactive data visualization function, which accepts excel file, X & Y column. 
    Chart types accepted are bar , scatter , pie , sun , histogram , box  , strip. 
    You can pass color column as well, having a boolean value.
    Image gets saved as .PNG in the same path as excel file.

    Usage: excel_charts(<excel path>,x_col='Name',y_col='Age', chart_type='bar',show_chart=True)
    """
    try:
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user('for data visualization')
            
        if not x_col:
            col_lst = excel_get_all_header_columns(excel_path, sheet_name, header)
            x_col = gui_get_dropdownlist_values_from_user('X Axis Column',col_lst,multi_select=False)[0]  

        if not y_col:
            col_lst = excel_get_all_header_columns(excel_path, sheet_name, header)
            y_col = gui_get_dropdownlist_values_from_user('Y Axis Column',col_lst,multi_select=False)[0]  

        if x_col and y_col:
            if color:
                df = pd.read_excel(excel_path,sheet_name=sheet_name,header=header,usecols={x_col,y_col,color})
            else:
                df = pd.read_excel(excel_path,sheet_name=sheet_name,header=header,usecols={x_col,y_col})

            fig = go.Figure()

            if chart_type == 'bar':

                fig.add_trace(go.Bar(x=df[x_col].values.tolist()))
                fig.add_trace(go.Bar(y=df[y_col].values.tolist()))

                if color:
                    fig = px.bar(df, x=x_col, y=y_col, barmode="group",color=color)
                else:
                    fig = px.bar(df, x=x_col, y=y_col, barmode="group")
                    
            elif chart_type == 'scatter':

                fig.add_trace(go.Scatter(x=df[x_col].values.tolist()))
                fig.add_trace(go.Scatter(y=df[x_col].values.tolist()))

            elif chart_type =='pie':

                if color:
                    fig = px.pie(df, names=x_col, values=y_col, title=title,color=color)#,hover_data=df.columns)
                else:
                    fig = px.pie(df, names=x_col, values=y_col, title=title)#,hover_data=df.columns)

            elif chart_type =='sun':

                if color:
                    fig = px.sunburst(df, path=[x_col], values=y_col,hover_data=df.columns,color=color)
                else:
                    fig = px.sunburst(df, path=[x_col], values=y_col,hover_data=df.columns)

            elif chart_type == 'histogram':

                if color:
                    fig = px.histogram(df, x=x_col, y=y_col, marginal="rug",color=color, hover_data=df.columns)
                else:
                    fig = px.histogram(df, x=x_col, y=y_col, marginal="rug",hover_data=df.columns)

            elif chart_type == 'box':

                if color:
                    fig = px.box(df, x=x_col, y=y_col, notched=True,color=color)
                else:
                    fig = px.box(df, x=x_col, y=y_col, notched=True)

            elif chart_type == 'strip':

                if color:
                    fig = px.strip(df, x=x_col, y=y_col, orientation="h",color=color)
                else:
                    fig = px.strip(df, x=x_col, y=y_col, orientation="h")

            fig.update_layout(title = title)
            
            if show_chart:
                fig.show()
            
            strFileName = _extract_filename_from_filepath(excel_path)
            strFileName = os.path.join(output_folder_path,strFileName + ".PNG")
            strFileName = Path(strFileName)
            
            scope = PlotlyScope()
            with open(strFileName, "wb") as f:
                f.write(scope.transform(fig, format="png"))
            print("Chart saved at " + str(strFileName))
        else:
            print("Please supply all the required values")

    except Exception as ex:
        print("Error in excel_draw_charts=" + str(ex))

def get_long_lat(strZipCode=0):
    """
    Function takes zip_code as input (int) and returns longitude, latitude, state, city, county. 
    """
    try:
        if not strZipCode:
            strZipCode = str(gui_get_any_input_from_user("USA Zip Code ex: 77429"))

        all_data_dict=zipcodes.matching(str(strZipCode))

        all_data_dict = all_data_dict[0]

        long = all_data_dict['long']
        lat = all_data_dict['lat']
        state = all_data_dict['state']
        city = all_data_dict['city']
        county = all_data_dict['county']
        return long, lat, state, city, county    
    except Exception as ex:
        print("Error in get_long_lat="+str(ex))

def excel_geotag_using_zipcodes(excel_path="",sheet_name='Sheet1',header=0,zoom_start=5,zip_code_column="",data_columns_as_list=[],color_boolean_column=""):
    """
    Function takes Excel file having ZipCode column as input. Takes one data column at present. 
    Creates .html file having geo-tagged markers/baloons on the page.

    Ex: excel_geotag_using_zipcodes()
    """

    try:
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user('for geo tagging (Note: As of now, works only for USA Zip codes)')

        if not zip_code_column:
            col_lst = excel_get_all_header_columns(excel_path, sheet_name, header)
            zip_code_column = gui_get_dropdownlist_values_from_user('having Zip Codes',col_lst,multi_select=False)[0]

        m = folium.Map(location=[40.178877,-100.914253 ], zoom_start=zoom_start)

        if len(data_columns_as_list) == 1:
            data_columns_as_str = str(data_columns_as_list).replace("[","").replace("]","").replace("'","")
        else:
            data_columns_as_str = str(data_columns_as_list).replace("[","").replace("]","")
            data_columns_as_str = data_columns_as_str[1:-1]
            
        use_cols = data_columns_as_list
        use_cols.append(zip_code_column)

        if color_boolean_column:
            use_cols.append(color_boolean_column)

        df = pd.read_excel(excel_path,sheet_name=sheet_name,header=header,usecols=use_cols)
        
        for _, row in df.iterrows():
            if not pd.isna(row[zip_code_column]) and str(row[zip_code_column]).isnumeric():
                
                long, lat, state, city, county = get_long_lat(str(row[zip_code_column]))
                county = str(county).replace("County","")
                
                if color_boolean_column and data_columns_as_str and row[color_boolean_column] == True:
                    folium.Marker(location=[lat, long], popup='State: ' + state + ',\nCity:' + city + ',\nCounty:' + county + ',\nDevice:' + row[data_columns_as_str], icon=folium.Icon(color='green', icon='info-sign')).add_to(m)
                elif data_columns_as_str:
                    folium.Marker(location=[lat, long], popup='State: ' + state + ',\nCity:' + city + ',\nCounty:' + county + ',\nDevice:' + row[data_columns_as_str], icon=folium.Icon(color='red', icon='info-sign')).add_to(m)
                else:
                    folium.Marker(location=[lat, long], popup='State: ' + state + ',\nCity:' + city + ',\nCounty:' + county, icon=folium.Icon(color='blue', icon='info-sign')).add_to(m)

        graphFileName = _extract_filename_from_filepath(excel_path)
        graphFileName = os.path.join(output_folder_path,graphFileName + ".html")
        graphFileName = Path(graphFileName)

        print("GeoTagged Graph saved at "+ graphFileName)
        m.save(graphFileName)
    
    except Exception as ex:
        print("Error in excel_geotag_using_zipcodes="+str(ex))
    
def _accept_cookies_h():
    """
    Internal function to accept cookies.
    """
    try:
        if Text('Accept cookies?').exists():
            click('I accept')
    except Exception as ex:
        print("Error in _accept_cookies_h="+str(ex))
    
def launch_website_h(URL="",dp=False,dn=True,igc=True,smcp=True,i=False,headless=False):
    try:
        """
        Internal function to launch browser.
        """
        if not URL:
            URL = gui_get_any_input_from_user("website URL to Launch Website using Helium functions. Ex https://www.google.com")

        global helium_service_launched
        helium_service_launched=True
        options = ChromeOptions()
        if dp:
            options.add_argument("--disable-popup-blocking")                
        if dn:  
            options.add_argument("--disable-notifications")                
        if igc:
            options.add_argument("--ignore-certificate-errors")             
        if smcp:
            options.add_argument("--suppress-message-center-popups")       
        if i:
            options.add_argument("--incognito")                             
        
        options.add_argument("--disable-translate")
        options.add_argument("--start-maximized")                          
        options.add_argument("--ignore-autocomplete-off-autofill")          
        options.add_argument("--no-first-run")                             
        #options.add_argument("--window-size=1920,1080")
        try:
            start_chrome(url=URL,options=options,headless=headless)
        except:
            try:
                start_firefox(url=URL,options=options,headless=headless) #to be tested
            except Exception as ex: 
                print('Either Chrome & Forefox is required='+str(ex))

        Config.implicit_wait_secs = 120
        _accept_cookies_h()
    except Exception as ex:
        print("Error in launch_website_h = "+str(ex))
        kill_browser()
        helium_service_launched = False
    
def browser_navigate_h(url="",dp=False,dn=True,igc=True,smcp=True,i=False,headless=False):
    try:
        """
        Navigates to Specified URL.
        """
        if not url:
            url = gui_get_any_input_from_user("website URL to Navigate using Helium functions. Ex: https://www.google.com")

        global helium_service_launched
        if not helium_service_launched:
            launch_website_h(URL=url,dp=dp,dn=dn,igc=igc,smcp=smcp,i=i,headless=headless)
            return
        go_to(url.lower())
        _accept_cookies_h()
    except Exception as ex:
        print("Error in browser_navigate_h = "+str(ex))
        helium_service_launched = False
    
def browser_write_h(Value="",User_Visible_Text_Element="",alert=False):
    """
    Write a string on the given element.
    """
    try:
        if not User_Visible_Text_Element:
            User_Visible_Text_Element = gui_get_any_input_from_user('visible element (placeholder) to WRITE your value. Ex: Username')

        if not Value:
            Value= gui_get_any_input_from_user('Value to be Written')

        if not alert:
            if Value and User_Visible_Text_Element:
                write(Value, into=User_Visible_Text_Element)
        if alert:
            if Value and User_Visible_Text_Element:
                write(Value, into=Alert(User_Visible_Text_Element))
    except Exception as ex:
        print("Error in browser_write_h = "+str(ex))
    
def browser_mouse_click_h(User_Visible_Text_Element="",element="d"):
    """
    click on the given element.
    """
    try:
        if not User_Visible_Text_Element:
            User_Visible_Text_Element = gui_get_any_input_from_user("visible text element (button/link/checkbox/radio etc) to Click")

        if User_Visible_Text_Element and element.lower()=="d":      #default
            click(User_Visible_Text_Element)
        elif User_Visible_Text_Element and element.lower()=="l":    #link
            click(link(User_Visible_Text_Element))
        elif User_Visible_Text_Element and element.lower()=="b":    #button
            click(Button(User_Visible_Text_Element))
        elif User_Visible_Text_Element and element.lower()=="t":    #textfield
            click(TextField(User_Visible_Text_Element))
        elif User_Visible_Text_Element and element.lower()=="c":    #checkbox
            click(CheckBox(User_Visible_Text_Element))
        elif User_Visible_Text_Element and element.lower()=="r":    #radiobutton
            click(RadioButton(User_Visible_Text_Element))
        elif User_Visible_Text_Element and element.lower()=="i":    #image ALT Text
            click(Image(alt=User_Visible_Text_Element))
    except Exception as ex:
        print("Error in browser_mouse_click_h = "+str(ex))
    
def browser_mouse_double_click_h(User_Visible_Text_Element=""):
    """
    Doubleclick on the given element.
    """
    try:
        if not User_Visible_Text_Element:
            User_Visible_Text_Element = gui_get_any_input_from_user("visible text element (button/link/checkbox/radio etc) to Double Click")

        if User_Visible_Text_Element:
            doubleclick(User_Visible_Text_Element)
    except Exception as ex:
        print("Error in browser_mouse_double_click_h = "+str(ex))
    
def browser_locate_element_h(element="",get_text=False):
    """
    Find the element by Xpath, id or css selection.
    """
    try:
        if not element:
            element = gui_get_any_input_from_user('browser element to locate (Helium)')
        if get_text:
            return S(element).web_element.text
        return S(element).web_element
    except Exception as ex:
        print("Error in browser_locate_element_h = "+str(ex))
    
def browser_locate_elements_h(element="",get_text=False):
    """
    Find the elements by Xpath, id or css selection.
    """
    try:
        if not element:
            element = gui_get_any_input_from_user('browser ElementS to locate (Helium)')
        if get_text:
            return find_all(S(element).web_element.text)
        return find_all(S(element).web_element)
    except Exception as ex:
        print("Error in browser_locate_elements_h = "+str(ex))
    
def browser_wait_until_h(text="",element="t"):
    """
    Wait until a specific element is found.
    """
    try:
        if not text:
            text = gui_get_any_input_from_user("visible text element to Search & Wait for")

        if element.lower()=="t":
            wait_until(Text(text).exists,10)        #text
        elif element.lower()=="b":
            wait_until(Button(text).exists,10)      #button
    except Exception as ex:
        print("Error in browser_wait_until_h = "+str(ex))

    
def browser_refresh_page_h():
    """
    Refresh the page.
    """
    try:
        refresh()
    except Exception as ex:
        print("Error in browser_refresh_page_h = "+str(ex))
    
def browser_hit_enter_h():
    """
    Hits enter KEY using Browser Helium Functions
    """
    try:
        press(ENTER)
    except Exception as ex:
        print("Error in browser_hit_enter_h="+str(ex))

def browser_quit_h():
    """
    Close the Helium browser.
    """
    try:
        kill_browser()
    except Exception as ex:
        print("Error in browser_quit_h = "+str(ex))

#Utility Functions
def dismantle_code(strFunctionName=""):
    """
    This functions dis-assembles given function and shows you column-by-column summary to explain the output of disassembled bytecode.

    Ex: dismantle_code(show_emoji)
    """
    try:
        if not strFunctionName:
            strFunctionName = gui_get_any_input_from_user('Exact function name to dis-assemble. Ex: show_emoji')
            print("Code dismantling {}".format(strFunctionName))
            return dis.dis(strFunctionName) 
    except Exception as ex:
       print("Error in dismantle_code="+str(ex)) 

def excel_clean_data(excel_path="",sheet_name='Sheet1',header=0,column_to_be_cleaned="",cleaning_pipe_line="Default"):
    """
    fillna(s) Replace not assigned values with empty spaces.
    lowercase(s) Lowercase all text.
    remove_digits() Remove all blocks of digits.
    remove_diacritics() Remove all accents from strings.
    remove_stopwords() Remove all stop words.
    remove_whitespace() Remove all white space between words.
    """
    try:
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user('to clean the data')
            
        if not column_to_be_cleaned:
            col_lst = excel_get_all_header_columns(excel_path, sheet_name, header)  
            column_to_be_cleaned = gui_get_dropdownlist_values_from_user('column list to Clean (removes digits/puntuation/stop words etc)',col_lst,multi_select=False)   
            column_to_be_cleaned = column_to_be_cleaned[0]

        if column_to_be_cleaned:
            df = pd.read_excel(excel_path,sheet_name=sheet_name,header=header)

            new_column_name = "Clean_" + column_to_be_cleaned

            if 'Default' in cleaning_pipe_line:
                df[new_column_name] = df[column_to_be_cleaned].pipe(hero.clean)
            else:
                custom_pipeline = [preprocessing.fillna, preprocessing.lowercase]
                df[new_column_name] = df[column_to_be_cleaned].pipe(hero.clean,custom_pipeline)    

            with pd.ExcelWriter(path=excel_path) as writer:
                df.to_excel(writer,index=False)

            print("Data Cleaned. Please see the output in {}".format(new_column_name))
    except Exception as ex:
        print("Error in excel_clean_data="+str(ex))
    
def compute_hash(inputData=""):
    """
    Returns the hash of the inputData 
    """
    try:
        if not inputData:
            inputData = gui_get_any_input_from_user('input string to compute Hash')

        return sha256(inputData.encode()).hexdigest()
    except Exception as ex:
        print("Error in compute_hash="+str(ex))

def browser_get_html_text(url=""):
    """
    Function to get HTML text without tags using Beautiful soup
    """
    try:
        if not url:
            url = gui_get_any_input_from_user("website URL to get HTML Text (without tags). Ex: https://www.cloint.com")

        html_text = requests.get(url) 
        soup = BeautifulSoup(html_text.content, 'lxml')
        text = str(soup.text).strip()
        text = ' '.join(text.split())
        return text
    except Exception as ex:
        print("Error in browser_get_html_text="+str(ex))

def word_cloud_from_url(url=""):
    """
    Function to create word cloud from a given website
    """
    try:
        text = browser_get_html_text(url=url)
        
        wc = WordCloud(max_words=2000, width=800, height=600,background_color='white',max_font_size=40, random_state=None, relative_scaling=0)
        wc.generate(text)
        file_path = os.path.join(output_folder_path,"URL_WordCloud.png")
        file_path = Path(file_path)

        wc.to_file(file_path)
        print("URL WordCloud saved at {}".format(file_path))

    except Exception as ex:
        print("Error in word_cloud_from_url="+str(ex))

def excel_describe_data(excel_path="",sheet_name='Sheet1',header=0):
    """
    Describe statistical data for the given excel
    """
    try:
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user("to Statistically Describe excel data")
            
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header)

        #user_option_lst = ['Numerical','String','Both']

        #user_choice = gui_get_dropdownlist_values_from_user("list of datatypes",user_option_lst)

        #if user_choice == 'Numerical':
        #    return df.describe(include = [np.number])
        #elif user_choice == 'String':
        #    return df.describe(include = ['O'])
        #else:
        #    return df.describe(include='all')
        return df.describe()

    except Exception as ex:
        print("Error in excel_describe_data="+str(ex))

def camera_capture_image(user_name=""):
    try:

        user_consent = gui_get_consent_from_user("turn ON camera & take photo ?")

        if user_consent == 'Yes':
            SECONDS = 5
            TIMER = int(SECONDS) 
            window_name = "ClointFusion"
            cap = cv2.VideoCapture(0) 

            if not cap.isOpened():
                print("Error in opening camera")

            cv2.namedWindow(window_name, cv2.WND_PROP_FULLSCREEN)
            cv2.setWindowProperty(window_name, cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)
            font = cv2.FONT_HERSHEY_SIMPLEX 

            if not user_name:
                user_name = gui_get_any_input_from_user("your name")

            while True: 

                ret, img = cap.read() 
                cv2.imshow(window_name, img) 
                prev = time.time() 

                text = "Taking selfie in 5 second(s)".format(str(TIMER))
                textsize = cv2.getTextSize(text, font, 1, 2)[0]
                print(str(textsize))

                textX = int((img.shape[1] - textsize[0]) / 2)
                textY = int((img.shape[0] + textsize[1]) / 2)

                while TIMER >= 0: 
                    ret, img = cap.read() 

                    cv2.putText(img, "Saving image in {} second(s)".format(str(TIMER)),  
                                (textX, textY ), font, 
                                1, (255, 0, 255), 
                                2) 
                    cv2.imshow(window_name, img) 
                    cv2.waitKey(125) 

                    cur = time.time() 

                    if cur-prev >= 1: 
                        prev = cur 
                        TIMER = TIMER-1

                ret, img = cap.read() 
                cv2.imshow(window_name, img) 
                cv2.waitKey(1000) 
                file_path = os.path.join(output_folder_path,user_name + ".PNG")
                file_path = Path(file_path)

                cv2.imwrite(file_path, img) 
                print("Image saved at {}".format(file_path))
                cap.release() 
                cv2.destroyAllWindows()
                break

        else:
            print("Operation cancelled by user")

    except Exception as ex:
        print("Error in camera_capture_image="+str(ex))   

          

def convert_csv_to_excel(csv_path="",sep=""):
    """
    Function to convert CSV to Excel 

    Ex: convert_csv_to_excel()
    """
    try:
        if not csv_path:
            csv_path = gui_get_any_file_from_user("CSV to convert to EXCEL","csv")

        if not sep:
            sep = gui_get_any_input_from_user("Delimeter Ex: |")

        csv_file_name = _extract_filename_from_filepath(csv_path)
        excel_file_name = csv_file_name + ".xlsx"        


        excel_file_path = os.path.join(output_folder_path,excel_file_name)
        excel_file_path = Path(excel_file_path)
        writer = pd.ExcelWriter(excel_file_path)

        df=pd.read_csv(csv_path,sep=sep)
        df.to_excel(writer, sheet_name='Sheet1', index=False)

        writer.save()
        
        print("Excel file saved : "+str(excel_file_path))

    except Exception as ex:
        print("Error in convert_csv_to_excel="+str(ex))





# Class related to capture_snip_now
class CaptureSnip(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        root = tk.Tk()
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        self.setGeometry(0, 0, screen_width, screen_height)
        self.setWindowTitle(' ')
        self.begin = QtCore.QPoint()
        self.end = QtCore.QPoint()
        self.setWindowOpacity(0.3)
        QtWidgets.QApplication.setOverrideCursor(
            QtGui.QCursor(QtCore.Qt.CrossCursor)
        )
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint)
        print('Capture now...')
        self.show()

    def paintEvent(self, event):
        qp = QtGui.QPainter(self)
        qp.setPen(QtGui.QPen(QtGui.QColor('black'), 3))
        qp.setBrush(QtGui.QColor(128, 128, 255, 128))
        qp.drawRect(QtCore.QRect(self.begin, self.end))

    def mousePressEvent(self, event):
        self.begin = event.pos()
        self.end = self.begin
        self.update()

    def mouseMoveEvent(self, event):
        self.end = event.pos()
        self.update()

    def mouseReleaseEvent(self, event):
        self.close()

        x1 = min(self.begin.x(), self.end.x())
        y1 = min(self.begin.y(), self.end.y())
        x2 = max(self.begin.x(), self.end.x())
        y2 = max(self.begin.y(), self.end.y())

        img = ImageGrab.grab(bbox=(x1, y1, x2, y2))
        file_num = str(len(os.listdir(img_folder_path)))
        file_name = os.path.join(img_folder_path,file_num + "_snip.PNG" )
        file_name = Path(file_name)

        print("Snip saved at " + str(file_name))
        img.save(file_name)
        
def capture_snip_now():
    """
    Captures the snip and stores in Image Folder of the BOT by giving continous numbering

    Ex: capture_snip_now()
    """
    app = ""
    try:
        if message_counter_down_timer(3):
            app = QtWidgets.QApplication(sys.argv)
            window = CaptureSnip()
            window.activateWindow()
            app.aboutToQuit.connect(app.deleteLater)
            sys.exit(app.exec_())
            
    except Exception as ex:
        print("Error in capture_snip_now="+str(ex))        
        try:
            sys.exit(app.exec_())
        except:
            pass

def ON_semi_automatic_mode():
    """
    This function sets semi_automatic_mode as True => ON
    """
    global enable_semi_automatic_mode
    semi_automatic_config_file_path = os.path.join(config_folder_path,"Semi_Automatic_Mode.txt")
    semi_automatic_config_file_path = Path(semi_automatic_config_file_path)

    try:    
        with open(semi_automatic_config_file_path, 'w') as f:
            f.write('True')
        enable_semi_automatic_mode = True
        print("Semi Automatic Mode is ENABLED "+ show_emoji())
    except Exception as ex:
        print("Error in ON_semi_automatic_mode="+str(ex))

def OFF_semi_automatic_mode():
    """
    This function sets semi_automatic_mode as False => OFF
    """
    global enable_semi_automatic_mode
    semi_automatic_config_file_path = os.path.join(config_folder_path,"Semi_Automatic_Mode.txt")
    semi_automatic_config_file_path = Path(semi_automatic_config_file_path)

    try:    
        with open(semi_automatic_config_file_path, 'w') as f:
            f.write('False')
        enable_semi_automatic_mode = False
        print("Semi Automatic Mode is DISABLED "+ show_emoji())
    except Exception as ex:
        print("Error in OFF_semi_automatic_mode="+str(ex))        

def _init_cf_quick_test_log_file(log_path_arg):
    """
    Internal function to generates the log and saves it to the file in the given base directory. 
    """
    global log_path
    log_path = log_path_arg
    from pif import get_public_ip

    try:
        dt_tm= str(datetime.datetime.now())    
        dt_tm = dt_tm.replace(" ","_")
        dt_tm = dt_tm.replace(":","-")
        dt_tm = dt_tm.split(".")[0]

        log_path = Path(os.path.join(log_path, str(dt_tm) + ".txt"))
                
        for handler in logging.root.handlers[:]:
            logging.root.removeHandler(handler)

        logging.basicConfig(filename=log_path, level=logging.INFO, format='%(asctime)s  :  %(message)s',datefmt='%Y-%m-%d %H:%M:%S')
        
    except Exception as ex:
        print("ERROR in _init_log_file="+str(ex))
    finally:
        host_ip = socket.gethostbyname(socket.gethostname()) 
        logging.info("{} ClointFusion Self Testing initiated".format(os_name))
        logging.info("{}/{}".format(host_ip,str(get_public_ip())))

def _download_cloint_quick_test_png():    
    """
    Internal function to download ClointFusion ICON from GitHub
    """
    try:
        if not os.path.exists(cf_icon_file_path):
            urllib.request.urlretrieve('https://raw.githubusercontent.com/ClointFusion/Image_ICONS_GIFs/main/Cloint-ICON.ico',cf_icon_file_path)

        if not os.path.exists(cf_logo_file_path):
            urllib.request.urlretrieve('https://raw.githubusercontent.com/ClointFusion/Image_ICONS_GIFs/main/Cloint-LOGO.PNG',cf_logo_file_path)
    except Exception as ex:
        print("Error while downloading Cloint ICOn/LOGO = "+str(ex))

def _rerun_clointfusion_first_run(ex):
    pg.alert("Please Re-run..."+str(ex))
    _,last_updated_date_file = is_execution_required_today('clointfusion_self_test',execution_type="M",save_todays_date_month=False)
    with open(last_updated_date_file, 'w',encoding="utf-8") as f:
        last_updated_on_date = int(datetime.date.today().strftime('%m')) - 1
        f.write(str(last_updated_on_date))

def clointfusion_self_test_cases(user_chosen_test_folder):
    """
    Main function for Self Test, which is called by GUI
    """
    global os_name
    chrome_close_PNG_1 = temp_current_working_dir / "Chrome-Close_1.PNG"
    chrome_close_PNG_2 = temp_current_working_dir / "Chrome-Close_2.PNG"
    chrome_close_PNG_3 = temp_current_working_dir / "Chrome-Close_3.PNG"

    twenty_PNG_1 = temp_current_working_dir / "Twenty_1.PNG"
    twenty_PNG_2 = temp_current_working_dir / "Twenty_2.PNG"
    twenty_PNG_3 = temp_current_working_dir / "Twenty_3.PNG"

    if not os.path.exists(chrome_close_PNG_1):
        urllib.request.urlretrieve('https://raw.githubusercontent.com/ClointFusion/Image_ICONS_GIFs/main/Chrome-Close_1.png',chrome_close_PNG_1)

    if not os.path.exists(chrome_close_PNG_2):
        urllib.request.urlretrieve('https://raw.githubusercontent.com/ClointFusion/Image_ICONS_GIFs/main/Chrome-Close_2.png',chrome_close_PNG_2)

    if not os.path.exists(chrome_close_PNG_3):
        urllib.request.urlretrieve('https://raw.githubusercontent.com/ClointFusion/Image_ICONS_GIFs/main/Chrome-Close_3.png',chrome_close_PNG_3)

    if not os.path.exists(twenty_PNG_1):
        urllib.request.urlretrieve('https://raw.githubusercontent.com/ClointFusion/Image_ICONS_GIFs/main/Twenty_1.png',twenty_PNG_1)

    if not os.path.exists(twenty_PNG_2):
        urllib.request.urlretrieve('https://raw.githubusercontent.com/ClointFusion/Image_ICONS_GIFs/main/Twenty_2.png',twenty_PNG_2)

    if not os.path.exists(twenty_PNG_3):
        urllib.request.urlretrieve('https://raw.githubusercontent.com/ClointFusion/Image_ICONS_GIFs/main/Twenty_3.png',twenty_PNG_3)

    test_folder_path = Path(os.path.join(user_chosen_test_folder,"ClointFusion_Self_Tests"))
    test_run_excel_path = Path(os.path.join(test_folder_path,'Quick_Self_Test_Excel.xlsx'))
    user_chosen_test_folder = Path(user_chosen_test_folder)
    test_folder_path = Path(test_folder_path)
    test_run_excel_path = Path(test_run_excel_path)

    try:
        message_pop_up('Importing ClointFusion')
        print('Importing ClointFusion')
        print()

        print('ClointFusion imported successfully '+ show_emoji())
        print("____________________________________________________________")
        print()
        logging.info('ClointFusion imported successfully')
        try:
            base_dir = Path(user_chosen_test_folder)
            folder_create(base_dir) 
            print('Test folder location {}'.format(base_dir))
            logging.info('Test folder location {}'.format(base_dir))
            
            img_folder_path =  os.path.join(base_dir, "Images")
            batch_file_path = os.path.join(base_dir, "Batch_File")
            config_folder_path = os.path.join(base_dir, "Config_Files")
            output_folder_path = os.path.join(base_dir, "Output")
            error_screen_shots_path = os.path.join(base_dir, "Error_Screenshots")
            
            try:
                print('Creating sub folders viz. img/batch/config/output/error_screen_shot at {}'.format(base_dir))
                folder_create(img_folder_path)
                folder_create(batch_file_path)
                folder_create(config_folder_path)
                folder_create(error_screen_shots_path)
                folder_create(output_folder_path)
            except Exception as ex:
                print('Unable to create basic sub-folders for img/batch/config/output/error_screen_shot=' + str(ex))
                logging.info('Unable to create basic sub-folders for img/batch/config/output/error_screen_shot')

            print()
            print('ClointFusion Self Testing Initiated '+show_emoji())
            logging.info('ClointFusion Self Testing Initiated')
        except Exception as ex:
            print('Error while creating sub-folders='+str(ex))
            logging.info('Error while creating sub-folders='+str(ex))

        try:
            print()
            print('Testing folder operations')
            folder_create(Path(os.path.join(test_folder_path,"My Test Folder")))
            folder_create_text_file(test_folder_path, "My Text File")
            excel_create_excel_file_in_given_folder(test_folder_path,'Quick_Self_Test_Excel')
            excel_create_excel_file_in_given_folder(test_folder_path,'My Excel-1')
            excel_create_excel_file_in_given_folder(test_folder_path,'My Excel-2')

            try:
                excel_create_excel_file_in_given_folder(os.path.join(test_folder_path,"Delete Excel"),'Delete-Excel-1')
                excel_create_excel_file_in_given_folder(os.path.join(test_folder_path,"Delete Excel"),'Delete-Excel-2')
                folder_delete_all_files(os.path.join(test_folder_path,'Delete Excel'), "xlsx")
            except Exception as ex:
                print('Unable to delete files in test folder='+str(ex))
                logging.info('Unable to delete files in test folder='+str(ex))

            folder_create(Path(test_folder_path / 'Split_Merge'))
            print(folder_get_all_filenames_as_list(test_folder_path))
            print(folder_get_all_filenames_as_list(test_folder_path, extension="xlsx"))
            print('Folder operations tested successfully '+show_emoji())
            print("____________________________________________________________")
            logging.info('Folder operations tested successfully')
        except Exception as ex:
            print('Error while testing Folder operations='+str(ex))
            logging.info('Error while testing Folder operations='+str(ex))

        if os_name == 'windows':
            try:
                print()
                print('Testing window based operations')
                window_show_desktop()
                launch_any_exe_bat_application(test_run_excel_path)
                window_minimize_windows('Quick_Self_Test_Excel')
                window_activate_and_maximize_windows('Quick_Self_Test_Excel')
                window_close_windows('Quick_Self_Test_Excel')
                print(window_get_all_opened_titles_windows())
                print('Window based operations tested successfully '+show_emoji())
                print("____________________________________________________________")
                logging.info('Window based operations tested successfully')
            except Exception as ex:
                print('Error while testing window based operations='+str(ex))
                logging.info('Error while testing window based operations='+str(ex))
        else:
            print('Skipping window operations as it is Windows OS specific')
            logging.info('Skipping window operations as it is Windows OS specific')

        try:
            print()
            print('Testing String Operations')
            print(string_remove_special_characters("C!@loin#$tFu*(sion"))
            print(string_extract_only_alphabets(inputString="C1l2o#%^int&*Fus12i5on"))
            print(string_extract_only_numbers("C1l2o3i4n5t6F7u8i9o0n"))
            print(date_convert_to_US_format("31-01-2021"))
            print('String operations tested successfully '+show_emoji())
            print("____________________________________________________________")
            logging.info('String operations tested successfully')
        except Exception as ex:
            print('Error while testing string operations='+str(ex))
            logging.info('Error while testing string operations='+str(ex))
        
        try:
            print()
            print('Testing keyboard operations')
            if os_name == 'windows':
                launch_any_exe_bat_application("notepad")
            else:
                launch_any_exe_bat_application("gedit") #Ubuntu / macOS ?

            if os_name == 'windows':
                key_write_enter("Performing ClointFusion Self Test for Notepad")
                key_hit_enter()
                key_press('alt+f4,n')
            else:
                pg.write("Performing ClointFusion Self Test for Text Editor / GEDIT")
                pg.press('enter')
                pg.hotkey('alt','f4')
                time.sleep(2)
                pg.hotkey('alt','w')
            
            message_counter_down_timer(3)
            
            print('Keyboard operations tested successfully '+show_emoji())
            print("____________________________________________________________")
            logging.info('Keyboard operations tested successfully')
        except Exception as ex:
            print('Error in keyboard operations='+str(ex))
            logging.info('Error in keyboard operations='+str(ex))
            try:
                key_press('alt+f4')
            except:
                pg.hotkey('alt','f4')

        message_counter_down_timer(3)
    
        try:
            print()
            print('Testing excel operations')
            excel_create_excel_file_in_given_folder(test_folder_path, "Test_Excel_File", "Test_Sheet")
            print(excel_get_row_column_count(test_run_excel_path))

            excel_create_excel_file_in_given_folder(test_folder_path,excelFileName="Excel_Test_Data")
            test_excel_path = test_folder_path / "Excel_Test_Data.xlsx"
            
            excel_set_single_cell(test_excel_path,columnName="Name",cellNumber=0,setText="A")
            excel_set_single_cell(test_excel_path,columnName="Name",cellNumber=1,setText="B")
            excel_set_single_cell(test_excel_path,columnName="Name",cellNumber=2,setText="C")
            excel_set_single_cell(test_excel_path,columnName="Name",cellNumber=3,setText="D")
            excel_set_single_cell(test_excel_path,columnName="Name",cellNumber=4,setText="E")
            excel_set_single_cell(test_excel_path,columnName="Age",cellNumber=0,setText="1")
            excel_set_single_cell(test_excel_path,columnName="Age",cellNumber=1,setText="2")
            excel_set_single_cell(test_excel_path,columnName="Age",cellNumber=2,setText="4")
            excel_set_single_cell(test_excel_path,columnName="Age",cellNumber=3,setText="3")
            excel_set_single_cell(test_excel_path,columnName="Age",cellNumber=4,setText="5")

            print(excel_get_single_cell(test_excel_path,sheet_name='Sheet1',columnName='Name'))

            excel_create_file(test_folder_path,"My New Paste Excel")

            excel_create_excel_file_in_given_folder(test_folder_path,'My Excel-3','CF-Sheet-1')
            excel_file_path = test_folder_path / 'My Excel-3.xlsx'
            print(excel_get_all_sheet_names(excel_file_path))
            print(excel_get_all_sheet_names(test_run_excel_path))
            
            excel_copied_Data=excel_copy_range_from_sheet(test_excel_path, sheet_name="Sheet1", startCol=1, startRow=1, endCol=2, endRow=6)
            print(excel_copied_Data)
            excel_copy_paste_range_from_to_sheet(Path(os.path.join(test_folder_path,"My New Paste Excel.xlsx")), sheet_name="Sheet1", startCol=1, startRow=1, endCol=2, endRow=6, copiedData=excel_copied_Data)
            excel_split_by_column(excel_path=Path(os.path.join(test_folder_path,"My New Paste Excel.xlsx")), sheet_name="Sheet1", header=0, columnName="Name")

            folder_create(Path(test_folder_path / 'Split_Merge'))
            excel_split_the_file_on_row_count(excel_path=Path(test_folder_path / "My New Paste Excel.xlsx"), sheet_name="Sheet1", rowSplitLimit=1, outputFolderPath=os.path.join(test_folder_path,'Split_Merge'), outputTemplateFileName="Split")
            excel_merge_all_files(input_folder_path=test_folder_path / "Split_Merge", output_folder_path=Path(test_folder_path,'Split_Merge'))
            excel_drop_columns(Path(test_folder_path / "My New Paste Excel.xlsx"), columnsToBeDropped ="Age")

            excel_sort_columns(excel_path=test_excel_path, sheet_name="Sheet1", header=0, firstColumnToBeSorted="Age", secondColumnToBeSorted="Name")
            excel_clear_sheet(Path(test_folder_path / "My New Paste Excel.xlsx"), sheet_name="Sheet1", header=0)

            excel_set_single_cell(test_excel_path,columnName="Name",cellNumber=5,setText="E")
            excel_set_single_cell(test_excel_path,columnName="Age",cellNumber=5,setText="5")
            excel_remove_duplicates(excel_path=test_excel_path, sheet_name="Sheet1", header=0,columnName="Name", which_one_to_keep="first")
            excel_create_file(test_folder_path,"My VLookUp Excel")
            excel_set_single_cell(Path(test_folder_path,"My VLookUp Excel.xlsx"),columnName="Name",cellNumber=0,setText="A")
            excel_set_single_cell(Path(test_folder_path,"My VLookUp Excel.xlsx"),columnName="Name",cellNumber=1,setText="B")
            excel_set_single_cell(Path(test_folder_path,"My VLookUp Excel.xlsx"),columnName="Name",cellNumber=2,setText="C")
            excel_set_single_cell(Path(test_folder_path,"My VLookUp Excel.xlsx"),columnName="Name",cellNumber=3,setText="D")
            excel_set_single_cell(Path(test_folder_path,"My VLookUp Excel.xlsx"),columnName="Name",cellNumber=4,setText="E")
            excel_set_single_cell(Path(test_folder_path,"My VLookUp Excel.xlsx"),columnName="Salary",cellNumber=0,setText="1")
            excel_set_single_cell(Path(test_folder_path,"My VLookUp Excel.xlsx"),columnName="Salary",cellNumber=1,setText="2")
            excel_set_single_cell(Path(test_folder_path,"My VLookUp Excel.xlsx"),columnName="Salary",cellNumber=2,setText="4")
            excel_set_single_cell(Path(test_folder_path,"My VLookUp Excel.xlsx"),columnName="Salary",cellNumber=3,setText="3")
            excel_set_single_cell(Path(test_folder_path,"My VLookUp Excel.xlsx"),columnName="Salary",cellNumber=4,setText="5")

            excel_vlook_up(filepath_1=test_excel_path,filepath_2=Path(test_folder_path,"My VLookUp Excel.xlsx"),match_column_name="Name")
            
            print('Excel operations tested successfully '+show_emoji())
            print("____________________________________________________________")
            logging.info('Excel operations tested successfully')
        except Exception as ex:
            print("Error while testing Excel Operations="+str(ex))
            logging.info("Error while testing Excel Operations="+str(ex))
            
        message_counter_down_timer(3)

        try:
            print()
            print("Testing screen-scraping functions")
            webbrowser.open('https://sites.google.com/view/clointfusion-hackathon') 
            message_counter_down_timer(5)
            
            pos=mouse_search_snip_return_coordinates_x_y(str(twenty_PNG_3),conf=0.5,wait=5)
            print(pos)
            pos=mouse_search_snip_return_coordinates_x_y(str(twenty_PNG_2),conf=0.5,wait=5)
            print(pos)
            pos=mouse_search_snip_return_coordinates_x_y(str(twenty_PNG_1),conf=0.5,wait=5)
            print(pos)

            pos=mouse_search_snips_return_coordinates_x_y([str(twenty_PNG_1),str(twenty_PNG_2),str(twenty_PNG_3)],conf=0.5,wait=10)
            print(pos)

            folder_create(os.path.join(test_folder_path,'Screen_scrape'))
            scrape_save_contents_to_notepad(test_folder_path / 'Screen_scrape')
                
            print("Screen-scraping functions tested successfully "+ show_emoji())
            print("____________________________________________________________")
            logging.info("Screen-scraping functions tested successfully")
        except Exception as ex:
            print('Error while testing screenscraping functions='+str(ex))
            logging.info('Error while testing screenscraping functions='+str(ex))

        try:
            print()
            print("Testing mouse operations")    
            mouse_move(850,600)
            print(mouse_get_color_by_position((800,500)))

            time.sleep(2)
            
            mouse_drag_from_to(600,510,1150,680)

            message_counter_down_timer(3)
            
            search_highlight_tab_enter_open("chat.whatsapp")

            pos = mouse_search_snips_return_coordinates_x_y([str(chrome_close_PNG_1),str(chrome_close_PNG_2),str(chrome_close_PNG_3)],conf=0.8,wait=3)
            print(pos)
            if pos is not None:
                mouse_click(*pos)

            pos = mouse_search_snips_return_coordinates_x_y([str(chrome_close_PNG_1),str(chrome_close_PNG_2),str(chrome_close_PNG_3)],conf=0.8,wait=3)
            print(pos)
            if pos is not None:
                mouse_click(*pos)

            mouse_click(int(pg.size()[0]/2),int(pg.size()[1]/2)) #Click at center of the screen

            print('Mouse operations tested successfully ' + show_emoji())
            print("____________________________________________________________")
            logging.info('Mouse operations tested successfully')
        except Exception as ex:
            print('Error in mouse operations='+str(ex))
            logging.info('Error in mouse operations='+str(ex))
            key_press('ctrl+w')
        
        message_counter_down_timer(3)

        try:
            print()
            print("Testing Browser's Helium functions")
            launch_website_h("https://pypi.org")
            browser_write_h("ClointFusion",User_Visible_Text_Element="Search projects")
            browser_hit_enter_h()
            browser_mouse_click_h("ClointFusion 0.")

            browser_mouse_double_click_h("RPA")
            
            browser_mouse_click_h("Open in Colab")
            
            browser_quit_h()
            print("Tested Browser's Helium functions successfully " + show_emoji())
            print("____________________________________________________________")
            logging.info("Tested Browser's Helium functions successfully")
        except Exception as ex:
            print("Error while Testing Browser Helium functions="+str(ex))
            logging.info("Error while Testing Browser Helium functions="+str(ex))
            key_press('ctrl+w') #to close any open browser

        message_counter_down_timer(3)
        
        try:
            print("____________________________________________________________")
            print()
            print("Congratulations - ClointFusion is compatible with your computer " + show_emoji('clap') + show_emoji('clap'))
            message_pop_up("Congratulations !!!\n\nClointFusion is compatible with your computer settings")
            
            logging.info("Flash message tested successfully")
        except Exception as ex:
            print("Error while testing Flash message="+str(ex))
            logging.info("Error while testing Flash message="+str(ex))


    except Exception as ex:
        print("ClointFusion Automated Testing Failed "+str(ex))
        logging.info("ClointFusion Automated Testing Failed "+str(ex))
        
    finally:
        _folder_write_text_file(Path(os.path.join(current_working_dir,'Running_ClointFusion_Self_Tests.txt')),str(False))
        print("____________________________________________________________")
        print("____________________________________________________________")
        print()
        print("ClointFusion Self Testing Completed")
        logging.info("ClointFusion Self Testing Completed")
        return True

def clointfusion_self_test():
    global os_name
    strEmail = ""
    start_time = time.monotonic()
    try:

        layout = [ [sg.Text("ClointFusion's First Run Setup",justification='c',font='Courier 18',text_color='orange')],
                [sg.T("Please enter your name",text_color='white'),sg.In(key='-NAME-',text_color='blue')],
                [sg.T("Please enter your email ID",text_color='white'),sg.In(key='-EMAIL-',text_color='blue')],
                [sg.T("I am ",text_color='white'),sg.Combo(values=['Student','Hobbyist','Professor','Professional','Others'], size=(35,30), key='-ROLE-',text_color='blue')],
                [sg.Text("We will be collecting & using ClointFusion's Self Test Report, to improve ClointFusion",justification='c',text_color='green',font='Courier 12')],
                [sg.Text('Its highly recommended to close all open files/folders/browsers before running this self test',justification='c',text_color='red',font='Courier 12')],
                [sg.Text('This Automated Self Test, takes around 4-5 minutes...Kindly do not move the mouse or type anything using keyboard.',justification='c',text_color='red',font='Courier 12')],
                [sg.Output(size=(140,20), key='-OUTPUT-')],
                [sg.Button('Start',bind_return_key=True,button_color=('white','green'),font='Courier 14'), sg.Button('Close',button_color=('white','firebrick'),font='Courier 14')]  ]

        if os_name == 'windows':
            window = sg.Window('Welcome to ClointFusion - Made in India with LOVE', layout, return_keyboard_events=True,use_default_focus=False,disable_minimize=True,grab_anywhere=False, disable_close=False,element_justification='c',keep_on_top=False,finalize=True,icon=cf_icon_file_path)
        else:
            window = sg.Window('Welcome to ClointFusion - Made in India with LOVE', layout, return_keyboard_events=True,use_default_focus=False,disable_minimize=False,grab_anywhere=False, disable_close=False,element_justification='c',keep_on_top=False,finalize=True,icon=cf_icon_file_path)
        
        while True:             
            event, values = window.read()

            if event == 'Start':
                try:
                    if values['-EMAIL-']:
                        valid = validate_email(str(values['-EMAIL-']))
                        strEmail = valid.email

                except EmailNotValidError as e:
                    pg.alert("Sorry, "+str(e))
                    
                if strEmail and values['-NAME-'] and values['-ROLE-']:
                
                    window['Start'].update(disabled=True)
                    window['Close'].update(disabled=True)
                    _folder_write_text_file(os.path.join(current_working_dir,'Running_ClointFusion_Self_Tests.txt'),str(True))

                    print("Starting ClointFusion's Automated Self Testing Module")
                    print('This may take several minutes to complete...')
                    print('During this test, some excel file, notepad, browser etc may be opened & closed automatically')
                    print('Please sitback & relax till all the test-cases are run...')
                    print()

                    _init_cf_quick_test_log_file(temp_current_working_dir)

                    if clointfusion_self_test_cases(temp_current_working_dir):
                        window['Close'].update(disabled=False)
                else:
                    pg.alert("Please enter all the values")

            if event in (sg.WIN_CLOSED, 'Close'):
                
                file_contents = ''
                try:
                    with open(log_path,encoding="utf-8") as f:
                        file_contents = f.readlines()
                except:
                    file_contents = 'Unable to read the file'

                if file_contents and file_contents != 'Unable to read the file':
                    from datetime import timedelta
                    time_taken= timedelta(seconds=time.monotonic()  - start_time)
                    
                    my_ip = "Hostname : {}".format(socket.gethostname()) + ", IP Address: " + str(socket.gethostbyname(socket.gethostname())) + " / " + str(get_public_ip())
                    my_id = values['-NAME-'] + ";" + strEmail + ";" + values['-ROLE-']
                    os_name = str(os_name) + ";" +  str(my_ip) + ";" +  str(my_id)

                    URL = 'https://docs.google.com/forms/d/e/1FAIpQLSehRuz_RWJDcqZMAWRPMOfV7CVZB7PjFruXZtQKXO1Q81jOgw/formResponse?usp=pp_url&entry.1012698071={}&entry.705740227={}&submit=Submit'.format(os_name + ";" + str(time_taken),file_contents)
                    webbrowser.open(URL)
                    message_counter_down_timer(10)
                    
                    #Ensure to close all browser if left open by this self test
                    time.sleep(2)
                    
                    try:
                        key_press('alt+f4')
                    except:
                        pg.hotkey('alt','f4')
                    time.sleep(2)
                    is_execution_required_today('clointfusion_self_test',execution_type="M",save_todays_date_month=True)
                break        
                    
    except Exception as ex:
        pg.alert('Error in Clointfusion Self Test = '+str(ex))
        _rerun_clointfusion_first_run(str(ex))
    finally:
        print('Thank you !')
        sys.exit(0)

# 4. All default services

# All new functions to be added before this line
# ########################
# ClointFusion's DEFAULT SERVICES

_welcome_to_clointfusion()

EXECUTE_SELF_TEST_NOW,last_updated_date_file = is_execution_required_today('clointfusion_self_test',execution_type="M")

if EXECUTE_SELF_TEST_NOW :
    try:
        my_id = "Hostname : {}".format(socket.gethostname()) + ", IP Address: " + str(socket.gethostbyname(socket.gethostname())) + " / " + str(get_public_ip())
        slack.post(text='Automated Self Test Initiated on {}'.format(my_id))
        clointfusion_self_test()
    except Exception as ex:
        print("Error in Self Test="+str(ex))
        _rerun_clointfusion_first_run(str(ex))
        
else:
    base_dir = gui_get_folder_path_from_user('Workspace Folder')

    if base_dir:
        base_dir = os.path.join(base_dir,"ClointFusion_BOT")
        base_dir = Path(base_dir)
        _set_bot_name()
        _download_cloint_ico_png()
        folder_create(base_dir) 

        log_path = Path(os.path.join(base_dir, "Logs"))
        img_folder_path =  Path(os.path.join(base_dir, "Images")) 
        batch_file_path = Path(os.path.join(base_dir, "Batch_File")) 
        config_folder_path = Path(os.path.join(base_dir, "Config_Files")) 
        output_folder_path = Path(os.path.join(base_dir, "Output")) 
        error_screen_shots_path = Path(os.path.join(base_dir, "Error_Screenshots"))
        status_log_excel_filepath = Path(os.path.join(base_dir,"StatusLogExcel"))

        folder_create(log_path)
        folder_create(img_folder_path)
        folder_create(batch_file_path)
        folder_create(config_folder_path)
        folder_create(error_screen_shots_path)
        folder_create(output_folder_path)
        _init_log_file()

        update_log_excel_file(bot_name +'- BOT initiated')
        _ask_user_semi_automatic_mode()
    else:
        pg.alert('Please re-run & select the Workspace Folder')
        sys.exit(0)

if os_name == 'windows':
    _load_missing_python_packages_windows()

# ########################

with warnings.catch_warnings():
    warnings.filterwarnings("ignore", category=PendingDeprecationWarning)
    warnings.filterwarnings("ignore", category=DeprecationWarning)