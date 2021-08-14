import os
import sys
import threading 
import time
import traceback
import platform
import subprocess
import sqlite3
import requests

from helium._impl import selenium_wrappers
from pyautogui import KEYBOARD_KEYS

cf_icon_file_path = "Cloint-ICON.ico"
cursr = ""
connct = ""
email = ""
passwd= ""

url = 'https://raw.githubusercontent.com/ClointFusion/ClointFusion/master/requirements.txt'

FIRST_TIME = False

current_working_dir = os.path.dirname(os.path.realpath(__file__)) #get cwd

os.chdir(current_working_dir)
try:
    os.system("{} -m pip install --upgrade pip".format(sys.executable))
except Exception as ex:
    print("Error updating PIP = " + str(ex) )

requirements_page = requests.get(url)
req_pkg_lst = str(requirements_page.text).splitlines()

req_pkg_lst = list(map(lambda s: s.strip(), req_pkg_lst))

def db_create_database_connect():
    """
    Function to create a database and connect to it
    """
    global cursr
    global connct
    try:            
        # connct = sqlite3.connect('{}.db'.format(database_name))
        connct = sqlite3.connect(r'{}\{}.db'.format(current_working_dir,"ClointFusion_DB"))
        cursr = connct.cursor()
        # print('Created & Connected with Database \'{}\''.format("ClointFusion_DB"))
    except Exception as ex:
        print("Error in db_create_database_connect="+str(ex))

def db_create_table():
    global cursr
    global connct
    try:
        table_name = 'My_Table'
        table_dict={'email': 'TEXT', 'passwd': 'TEXT'}
        table = str(table_dict).replace("{","").replace("'","").replace(":","").replace("}","")
        # table = table.replace('INT,','INT PRIMARY KEY,',1) #make first field as PK
        exec_query = "CREATE TABLE IF NOT EXISTS {}({});".format(table_name,table)
        cursr.execute("""{}""".format(exec_query))
        connct.commit()
        # print('Table \'{}\' created'.format(table_name))
    except Exception as ex:
        print("Error in db_create_table="+str(ex))

def db_check_record():
    global cursr
    global connct
    global email, passwd

    table_name = 'My_Table'
    exec_query = "SELECT * FROM {};".format(table_name) 
    cursr.execute(exec_query)
    all_results = cursr.fetchall()

    if all_results:
        email = all_results[0][0]
        passwd = all_results[0][1]
    return all_results

def db_insert_rows(email, passwd):
    global cursr
    global connct

    table_name = 'My_Table'
    table_dict = {'email':email,'passwd':passwd}

    table_keys = str(table_dict.keys()).replace('dict_keys([',"").replace("'","").replace("])","")
    table_values = str(table_dict.values()).replace('dict_values([',"").replace("])","")

    exec_query = "INSERT INTO {}({}) VALUES({});".format(table_name,table_keys,table_values)
    
    cursr.execute("""{}""".format(exec_query))
    connct.commit()
    # print("Row with values {} inserted into \'{}\'".format(table_values,table_name))
    
def _load_missing_python_packages_windows(list_of_required_packages_1=[]):
    """
    Installs Windows OS specific python packages
    """       
    try:
        list_of_required_packages = [x.strip().lower() for x in list_of_required_packages_1]
        reqs = subprocess.check_output([sys.executable, '-m', 'pip', 'list'])
        installed_packages = [str(r.decode().split('==')[0]).strip().lower() for r in reqs.split()]
        
        missing_packages = ' '.join(list(set(list_of_required_packages)-set(installed_packages)))
        if missing_packages:
            print("{} package(s) are missing".format(missing_packages)) 
            
            if "comtypes" in missing_packages:
                os.system("{} -m pip install comtypes==1.1.7".format(sys.executable))
            
            for pkg in missing_packages:
                pkg_with_version = filter(lambda a: pkg in a, req_pkg_lst)
                # print(pkg_with_version)
                cmd = "pip install {}".format(list(pkg_with_version)[0])
            # print(cmd)
            os.system(cmd) 

    except Exception as ex:
        print("Error in _load_missing_python_packages_windows="+str(ex))

try:
    import pyautogui as pg
except Exception as ex:
    _load_missing_python_packages_windows(['pyautogui'])
    import pyautogui as pg
    
os_name = str(platform.system()).lower()

if os_name != 'windows':
    pg.alert("Colab Launcher works only on windows OS as of now")
    exit(0)

try:
    import psutil
except:
    _load_missing_python_packages_windows(["psutil"])
    import psutil

def is_chrome_open():    
    try:
        for proc in psutil.process_iter(['pid', 'name']):
            # This will check if there exists any process running with executable name
            if proc.info['name'] == 'chrome.exe':
                yes_no=pg.confirm(text='Chrome browser needs to be closed !\n\nPlease click "Yes" to forcefully close it', title="ClointFusion's Colab Launcher", buttons=['Yes', 'No'])
                
                if yes_no == 'Yes':
                    try:
                        subprocess.call("TASKKILL /f /IM CHROME.EXE",stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)
                    except:
                        pass
                    try:
                        subprocess.call("TASKKILL /f /IM CHROMEDRIVER.EXE",stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)
                    except:
                        pass
                    return False
                else:
                    return True

    except Exception as ex:
        pg.alert("Error while closing chrome")
        exc_type, exc_value, exc_tb = sys.exc_info()
        pg.alert(traceback.format_exception(exc_type, exc_value, exc_tb,limit=None, chain=True))
        exit(0)

if is_chrome_open()==True:
    pg.alert("Please close Google Chrome browser & try again")
    exit(0)
    
# try:
    # os.system("pip install -r {}".format(requirements_path))
# except Exception as ex:
try:
    _load_missing_python_packages_windows(['setuptools ','wheel', 'watchdog','Pillow','pynput','pif','PyAutoGUI ','PySimpleGUI ','bs4','clipboard','emoji','folium ','helium','imutils','kaleido','keyboard','matplotlib','numpy','opencv-python','openpyxl','pandas','plotly','requests','selenium','texthero','wordcloud','zipcodes','pathlib3x','pathlib','PyQt5','email-validator','testresources','scikit-image ','pivottablejs','ipython ','comtypes','cryptocode','ImageHash','get-mac','xlsx2html ','simplegmail','xlwings ','jupyterlab','notebook','Pygments','psutil','gspread'])    
except Exception as ex:
    pg.alert("Error while executing pip install -r requirements.txt")
    exc_type, exc_value, exc_tb = sys.exc_info()
    pg.alert(traceback.format_exception(exc_type, exc_value, exc_tb,limit=None, chain=True))

# try:
#     import sourcedefender
# except Exception as ex:
#     _load_missing_python_packages_windows(['sourcedefender'])
#     print(str(ex))
#     import sourcedefender
    
# finally:
#     import ClointFusion_Lite as cfl
    
try:
    import chromedriver_binary
except:
    _load_missing_python_packages_windows(['chromedriver-binary-auto'])
    import chromedriver_binary

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from helium import *
except:
    _load_missing_python_packages_windows(['selenium','helium'])
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from helium import *

# try:
#     import ClointFusion
# except Exception as ex:
#     try:
#         # os.system("pip install ClointFusion")
#         _load_missing_python_packages_windows(['clointfusion'])
#     except:
#         pg.alert("Error while executing pip install ClointFusion")    
#         exc_type, exc_value, exc_tb = sys.exc_info()
#         pg.alert(traceback.format_exception(exc_type, exc_value, exc_tb,limit=None, chain=True))
#         sys.exit(0)

try:
    import keyboard as kb
    import PySimpleGUI as sg
    import pygetwindow as gw
    sg.theme('Dark') # for PySimpleGUI FRONT END        
except:
    _load_missing_python_packages_windows(['keyboard','PySimpleGUI','PyGetWindow'])
    import keyboard as kb
    import PySimpleGUI as sg
    import pygetwindow as gw
    sg.theme('Dark') # for PySimpleGUI FRONT END        

def launch_jupyter(): 
    try:
        cmd = "pip install --upgrade jupyter_http_over_ws>=0.0.7 && jupyter serverextension enable --py jupyter_http_over_ws"
        # subprocess.call(cmd,stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)
        os.system(cmd)

        cmd = 'jupyter notebook --no-browser --allow-root --NotebookApp.allow_origin="https://colab.research.google.com" --NotebookApp.token=""  --NotebookApp.disable_check_xsrf=True'
        # subprocess.call(cmd,stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)
        os.system(cmd)
    except:
        print("Error in launch_jupyter")
        pg.alert("Error in launch_jupyter")

        #Kill the port if busy
        try:
            os.system('taskkill /F /PID  8888')

            cmd = "pip install --upgrade jupyter_http_over_ws>=0.0.7 && jupyter serverextension enable --py jupyter_http_over_ws"
            os.system(cmd)

            cmd = 'jupyter notebook --no-browser --allow-root --NotebookApp.allow_origin="https://colab.research.google.com" --NotebookApp.token=""  --NotebookApp.disable_check_xsrf=True'
                #   'jupyter notebook --NotebookApp.allow_origin='https://colab.research.google.com' --NotebookApp.port_retries=0 --notebook-dir="" --no-browser --allow-root --NotebookApp.token='' --NotebookApp.disable_check_xsrf=True --port=8888
            os.system(cmd)
        except Exception as ex:
            print("Port is busy = "+str(ex))

db_create_database_connect()
db_create_table()

def get_email_password_from_user():
    global FIRST_TIME
    try:

        layout = [[sg.Text("ClointFusion - Set Yourself Free for Better Work", font='Courier 16', text_color='orange')],
                [sg.Text(text='Please enter Gmail ID:',font=('Courier 12'),text_color='yellow'),sg.Input(key='-GMAIL-', justification='c',focus=True)],
                [sg.Text(text='Please enter Password:',font=('Courier 12'),text_color='yellow'),sg.Input(key='-PASSWD-', justification='c',password_char='*')],
                [sg.Submit('OK',button_color=('white','green'),bind_return_key=True, focus=True),sg.CloseButton('Cancel',button_color=('white','firebrick'))],
                [sg.Text("These credentials will be stored on you local computer, used to automatically login & will be associated with Colab Launcher")]]

        window = sg.Window('ClointFusion - Colab Launcher',layout, return_keyboard_events=True,use_default_focus=False,disable_close=False,element_justification='c',keep_on_top=True, finalize=True,icon=cf_icon_file_path)

        while True:                
            event, values = window.read()

            if event is None or event == 'Cancel' or event == "Escape:27":
                values = []
                # break
                sys.exit(0)

            if event == 'OK':
                if values and values['-GMAIL-'] and values['-PASSWD-']:
                    db_insert_rows(values['-GMAIL-'],values['-PASSWD-'])
                    FIRST_TIME = True
                    break
                else:
                    pg.alert("Please enter all the values")

        window.close()
        
    except Exception as ex:
        print("Error in get_colab_url_from_user="+str(ex))

def db_delete_data():
    global cursr

    cursr.execute("""{}""".format("DELETE FROM 'My_Table' WHERE email='mayur@cloint.com'"))
    all_results = cursr.fetchall()
    print(all_results)

# db_delete_data()

if not db_check_record():
    get_email_password_from_user()

def get_colab_url_from_user():
    ret_val = "cancelled"
    try:
        dropdown_list = ["ClointFusion Labs (Public)", "ClointFusion Starter (Hackathon)"] #"ClointFusion Lite (Interns Only)"
        oldKey = "Please choose desired Colab :"
        # oldValue = "https://colab.research.google.com/github/ClointFusion/ClointFusion/blob/master/ClointFusion_Labs.ipynb"
        oldValue = 'ClointFusion Labs (Public)'

        layout = [[sg.Text("ClointFusion - Set Yourself Free for Better Work", font='Courier 16', text_color='orange')],
                [sg.Text(text=oldKey,font=('Courier 12'),text_color='yellow'),sg.Listbox(dropdown_list,size=(30, 5),key='user_choice',default_values=oldValue,enable_events=True,change_submits=True)],#oluser_choice
                [sg.Submit('OK',button_color=('white','green'),bind_return_key=True, focus=True),sg.CloseButton('Cancel',button_color=('white','firebrick'))],
                [sg.Text("This is an automated tool which connects ClointFusion Colab with your Local Runtime.\nSign-in using your Gmail ID & wait for setup to Finish..")]]

        window = sg.Window('ClointFusion - Colab Launcher',layout, return_keyboard_events=True,use_default_focus=False,disable_close=False,element_justification='c',keep_on_top=True, finalize=True,icon=cf_icon_file_path)

        while True:                
            event, values = window.read()

            if event is None or event == 'Cancel' or event == "Escape:27":
                values = []
                break

            if event == 'OK':
                if values and values['user_choice']:
                    ret_val = str(values['user_choice'][0])
                    break
                else:
                    pg.alert("Please enter all the values")

        window.close()
        
    except Exception as ex:
        print("Error in get_colab_url_from_user="+str(ex))
        
    finally:
        return ret_val

def modify_file_as_text(text_file_path, text_to_search, replacement_text):
    import fileinput

    with fileinput.FileInput(text_file_path, inplace=True, backup='.bak') as file:
        for line in file:
            print(line.replace(text_to_search, replacement_text), end='')

def connect_to_local_runtime(user_choice): 

    try:
        
        if user_choice == "ClointFusion Labs (Public)":
            colab_url = "https://accounts.google.com/signin/v2/identifier?authuser=0&hl=en&continue=https://colab.research.google.com/github/ClointFusion/ClointFusion/blob/master/ClointFusion_Labs.ipynb" #https://colab.research.google.com/github/ClointFusion/ClointFusion/blob/master/ClointFusion_Labs.ipynb"
            # colab_url = "https://colab.research.google.com/github/ClointFusion/ClointFusion/blob/master/ClointFusion_Labs.ipynb"
                        
        # elif user_choice == "ClointFusion Lite (Interns Only)":
        #     #Extract encrypted version of ClointFusion_Lite to a specific folder and in Colab import that folder
        #     colab_url = 'https://accounts.google.com/signin/v2/identifier?authuser=0&hl=en&continue=https://colab.research.google.com/drive/11MvoQfNFXJqlXKcXV1LBVUE98Ks48M_a'

        elif user_choice == "ClointFusion Starter (Hackathon)":        
            colab_url = 'https://accounts.google.com/signin/v2/identifier?authuser=0&hl=en&continue=https://colab.research.google.com/drive/1G9mh58z8AbWqBit2TC4Wgg6p_eHPvUJB'

        user_data_path = "C:\\Users\\{}\\AppData\\Local\\Google\\Chrome\\User Data".format(os.getlogin())

        modify_file_as_text(user_data_path + '\\Default\\Preferences', 'crashed', 'false')

        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--user-data-dir=' + user_data_path)
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument('--ignore-certificate-errors')
        chrome_options.add_argument("--disable-session-crashed-bubble")
        chrome_options.add_argument("--disable-popup-blocking")
        chrome_options.add_argument("--disable-notifications")
        chrome_options.add_argument("--suppress-message-center-popups")
        chrome_options.add_argument("--disable-translate")
        chrome_options.add_argument("--no-first-run") 
        chrome_options.add_argument("--disable-extensions") 
        chrome_options.add_argument("--disable-application-cache") 

        driver = start_chrome(options=chrome_options, url=colab_url)

        # kb.press_and_release('win+d')

        chrome = gw.getWindowsWithTitle('Google Chrome')[0]
        chrome.activate()

        # pg.doubleClick(pg.size()[0]/2,pg.size()[1]/2)
        # kb.press_and_release('esc')
        # kb.press_and_release('esc')

        try:
            wait_until(Text("Code").exists,timeout_secs=6)
            
        except :#selenium_wrappers.common.exceptions.TimeoutException:
            try:
                click(email)
            except:
                write(email, into='Email or phone')

            click('Next')
            time.sleep(0.5)
            write(passwd, into='Enter your password')
            click('Next')
            time.sleep(0.5)

            wait_until(Text("Code").exists,timeout_secs=240)

        # kb.press_and_release('esc')
        # time.sleep(0.2)

        # pg.press(ESCAPE)
        # time.sleep(0.2)

        # press(ESCAPE)
        # time.sleep(0.2)

        if FIRST_TIME:
            #create short-cut
            press(CONTROL + 'mh')
            time.sleep(1)

            v = S("//input[@id='pref_shortcut_connectLocal']")

            write('',v)

            press(CONTROL + '1')
            time.sleep(0.5)
            
            click("SAVE")

            time.sleep(1)

        #use short-cut

        press(CONTROL + '1')
        time.sleep(1)
        # pg.alert("HKHR")

        pg.doubleClick(pg.size()[0]/2,pg.size()[1]/2)
        time.sleep(1)

        if FIRST_TIME:
            
            kb.press_and_release('SHIFT+TAB')
            time.sleep(0.5)
            kb.press_and_release('SHIFT+TAB')
            time.sleep(0.5)
            kb.press_and_release('SHIFT+TAB')
            time.sleep(0.5)

            kb.write("http://localhost:8888")
            time.sleep(2)

        # click("CONNECT")
            kb.press_and_release('TAB')
            time.sleep(0.5)
            # pg.alert(1)
            kb.press_and_release('TAB')
            time.sleep(0.5)
            # pg.alert(2)

        else:
            kb.press_and_release('SHIFT+TAB')
            time.sleep(0.5)

        press(ENTER)
        time.sleep(2)

        # try:
        #     img = "Restore_Bubble.PNG"
        #     pos = pg.locateOnScreen(img, confidence=0.8)  #region=
        #     pg.alert(pos)
        #     pg.click(*pos)
        # except:
        #     pass

        pg.alert("Ready ! Google Colab is now connected with your Local Runtime.\n\nPlease click 'OK' & you are all set to work on ClointFusion Colabs...")

    except Exception as ex:
        print("Error in connect_to_local_runtime="+str(ex))
        exc_type, exc_value, exc_tb = sys.exc_info()
        pg.alert(traceback.format_exception(exc_type, exc_value, exc_tb,limit=None, chain=True))

        pg.alert("Error in connect_to_local_runtime="+str(ex))
        connect_to_local_runtime()

# def popup_msg():
#     sg.PopupTimed("Loading... Please wait", auto_close=30)

if __name__ == "__main__": 
    try:
        user_choice = get_colab_url_from_user()
        
        if user_choice != "cancelled":

            # creating threads 
            t1 = threading.Thread(target=connect_to_local_runtime,args=(user_choice,))
            t2 = threading.Thread(target=launch_jupyter) 
            # t3 = threading.Thread(target=popup_msg)
            
            t1.start() 
            t2.start() 
            # t3.start()

            t1.join() 
            t2.join() 
            # t3.join()
        else:
            print("User Cancelled the Launch")

    except Exception as ex:

        pg.alert("Error in Main="+str(ex))

        exc_type, exc_value, exc_tb = sys.exc_info()
        pg.alert(traceback.format_exception(exc_type, exc_value, exc_tb,limit=None, chain=True))
        print("Error in Main="+str(ex))
        