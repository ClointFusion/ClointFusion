import subprocess
import os
import sys
import platform
import sqlite3
import time
import datetime
import traceback
import shutil
import socket
import threading
import requests
import webbrowser
from dateutil import parser
from datetime import timedelta
import PySimpleGUI as sg
import openpyxl as op
from dateutil import parser
from pathlib import Path
import pyautogui as pg
import pandas as pd
import tempfile
import random
import pyinspect as pi
from rich import pretty

temp_current_working_dir = tempfile.mkdtemp(prefix="cloint_",suffix="_fusion")
temp_current_working_dir = Path(temp_current_working_dir)

windows_os = "windows"
linux_os = "linux"
mac_os = "darwin"
os_name = str(platform.system()).lower()

python_exe_path = os.path.join(os.path.dirname(sys.executable), "python.exe")
pythonw_exe_path = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")

if os_name == windows_os:
    clointfusion_directory = r"C:\Users\{}\ClointFusion".format(str(os.getlogin()))
elif os_name == linux_os:
    clointfusion_directory = r"/home/{}/ClointFusion".format(str(os.getlogin()))
elif os_name == mac_os:
    clointfusion_directory = r"/Users/{}/ClointFusion".format(str(os.getlogin()))
else:
    clointfusion_directory = temp_current_working_dir

current_working_dir = os.path.dirname(os.path.realpath(__file__)) #get cwd
config_folder_path = Path(os.path.join(clointfusion_directory, "Config_Files"))
img_folder_path =  Path(os.path.join(clointfusion_directory, "Images"))
cf_splash_png_path = Path(os.path.join(clointfusion_directory,"Logo_Icons","Splash.PNG"))
cf_icon_cdt_file_path = os.path.join(clointfusion_directory,"Logo_Icons","Cloint-ICON-CDT.ico")
log_path = Path(os.path.join(clointfusion_directory, "Logs"))
batch_file_path = Path(os.path.join(clointfusion_directory, "Batch_File"))    
output_folder_path = Path(os.path.join(clointfusion_directory, "Output")) 
error_screen_shots_path = Path(os.path.join(clointfusion_directory, "Error_Screenshots"))
status_log_excel_filepath = Path(os.path.join(clointfusion_directory,"StatusLogExcel"))
cf_icon_file_path = os.path.join(clointfusion_directory,"Logo_Icons","Cloint-ICON.ico")
cf_logo_file_path = os.path.join(clointfusion_directory,"Logo_Icons","Cloint-LOGO.PNG")


try:
    db_file_path = r'{}\BRE_WHM.db'.format(str(config_folder_path))
    connct = sqlite3.connect(db_file_path,check_same_thread=False)
    cursr = connct.cursor()
except Exception as ex:
    print("Error in connecting to DB="+str(ex))

def _get_site_packages_path():
    """
    Returns Site-Packages Path
    """
    import subprocess
    try:
        import site  
        site_packages_path = next(p for p in site.getsitepackages() if 'site-packages' in p)
    except:
        site_packages_path = subprocess.run('python -c "import os; print(os.path.join(os.path.dirname(os.__file__), \'site-packages\'))"',capture_output=True, text=True).stdout

    site_packages_path = str(site_packages_path).strip()  
    return str(site_packages_path)

def call_social_media():
    #opens all social media links of ClointFusion
    try:
        webbrowser.open_new_tab("https://www.facebook.com/ClointFusion")
    except Exception as ex:
        print("Error in call_social_media = " + str(ex))
 
    try:
        webbrowser.open_new_tab("https://twitter.com/ClointFusion")
    except Exception as ex:
        print("Error in call_social_media = " + str(ex))

    try:
        webbrowser.open_new_tab("https://www.youtube.com/channel/UCIygBtp1y_XEnC71znWEW2w")
    except Exception as ex:
        print("Error in call_social_media = " + str(ex))

    try:
        webbrowser.open_new_tab("https://www.linkedin.com/showcase/clointfusion_official")
    except Exception as ex:
        print("Error in call_social_media = " + str(ex))
    try:
        webbrowser.open_new_tab("https://www.reddit.com/user/Cloint-Fusion")
    except Exception as ex:
        print("Error in call_social_media = " + str(ex))

    try:
        webbrowser.open_new_tab("https://www.instagram.com/clointfusion")
    except Exception as ex:
        print("Error in call_social_media = " + str(ex))

    try:
        webbrowser.open_new_tab("https://www.kooapp.com/profile/ClointFusion")
    except Exception as ex:
        print("Error in call_social_media = " + str(ex))

    try:
        webbrowser.open_new_tab("https://discord.com/invite/tsMBN4PXKH")
    except Exception as ex:
        print("Error in call_social_media = " + str(ex))

    try:
        webbrowser.open_new_tab("https://www.eventbrite.com/e/2-days-event-on-software-bot-rpa-development-with-no-coding-tickets-183070046437")
    except Exception as ex:
        print("Error in call_social_media = " + str(ex))

    try:
        webbrowser.open_new_tab("https://internshala.com/internship/detail/python-rpa-automation-software-bot-development-work-from-home-job-internship-at-clointfusion1631715670")
    except Exception as ex:
        print("Error in call_social_media = " + str(ex))

site_packages_path = _get_site_packages_path()