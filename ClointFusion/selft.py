import uuid, socket, os, webbrowser, requests, platform, subprocess, sqlite3, tempfile, logging, sys, time
from pathlib import Path
import helium as browser
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import random, threading, urllib
import traceback

os_name = str(platform.system()).lower()
windows_os = "windows"
linux_os = "linux"
mac_os = "darwin"
FIRST_RUN = ""
SELF_TEST = ""
temp_current_working_dir = tempfile.mkdtemp(prefix="cloint_",suffix="_fusion")
temp_current_working_dir = Path(temp_current_working_dir)

connct, cursr = "", ""
c_version = ""
s_version = ""
need_self_test = ""
username = os.getlogin()
email = ""
# Paths to the files and folders
if os_name == windows_os:
    system_uuid = str(subprocess.check_output('wmic csproduct get uuid'), 'utf-8').split('\n')[1].strip()
    clointfusion_directory = r"C:\Users\{}\ClointFusion".format(str(os.getlogin()))
elif os_name == linux_os:
    system_uuid = str(subprocess.check_output('sudo dmidecode -s system-uuid', shell=True),'utf-8').split('\n')[0].strip()
    clointfusion_directory = r"/home/{}/ClointFusion".format(str(os.getlogin()))
elif os_name == mac_os:
    system_uuid = str(uuid.uuid5(uuid.NAMESPACE_URL, socket.gethostname())).upper()
    clointfusion_directory = r"/Users/{}/ClointFusion".format(str(os.getlogin())) 
else:
    clointfusion_directory = temp_current_working_dir

config_folder_path = Path(os.path.join(clointfusion_directory, "Config_Files"))
log_path = Path(os.path.join(clointfusion_directory, "Logs"))
img_folder_path =  Path(os.path.join(clointfusion_directory, "Images")) 
batch_file_path = Path(os.path.join(clointfusion_directory, "Batch_File"))    
output_folder_path = Path(os.path.join(clointfusion_directory, "Output")) 
error_screen_shots_path = Path(os.path.join(clointfusion_directory, "Error_Screenshots"))
status_log_excel_filepath = Path(os.path.join(clointfusion_directory,"StatusLogExcel"))


cf_icon_file_path = os.path.join(clointfusion_directory,"Logo_Icons","Cloint-ICON.ico")
cf_icon_cdt_file_path = os.path.join(clointfusion_directory,"Logo_Icons","Cloint-ICON-CDT.ico")
cf_logo_file_path = os.path.join(clointfusion_directory,"Logo_Icons","Cloint-LOGO.PNG")
cf_splash_png_path = Path(os.path.join(clointfusion_directory,"Logo_Icons","Splash.PNG"))

db_file_path = r'{}\BRE_WHM.db'.format(str(config_folder_path))
update_path = r'{}\update.pyw'.format(str(config_folder_path))

# Private Functions
def _get_site_packages_path():
    """
    Returns Site-Packages Path
    """
    try:
        import site  
        site_packages_path = next(p for p in site.getsitepackages() if 'site-packages' in p)
    except:
        site_packages_path = subprocess.run('python -c "import os; print(os.path.join(os.path.dirname(os.__file__), \'site-packages\'))"',capture_output=True, text=True).stdout

    site_packages_path = str(site_packages_path).strip()  
    
    return str(site_packages_path)

def _self_test():
    global cursr
    data = cursr.execute("SELECT self_test from CF_VALUES")
    for row in data:
        if row[0] == "True":
            return True
        else:
            return False

def _update_registry(file_path):
    """
    Add BRE_WHM to Registry for AutoRun
    """
    try:
        import winreg as reg 
        address=file_path
        # key we want to change is HKEY_CURRENT_USER 
        # key value is Software\Microsoft\Windows\CurrentVersion\Run

        key_value = "Software\Microsoft\Windows\CurrentVersion\Run"
        open = reg.OpenKey(reg.HKEY_CURRENT_USER,key_value,0,reg.KEY_ALL_ACCESS)

        # key = reg.OpenKey(reg.HKEY_CURRENT_USER, key_value, 0, reg.KEY_ALL_ACCESS)        
        # reg.SetValueEx(open,"ClointFusion",0,reg.REG_SZ,address)

        reg.DeleteValue(open, 'ClointFusion')
        reg.CloseKey(open)
    except Exception as ex:
        pass

def _create_short_cut(short_cut_path="",target_file_path="",work_dir=""):
    if os_name == windows_os:
        try:    
            import winshell
            from win32com.client import Dispatch

            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(short_cut_path) # Path to be saved (shortcut)
            shortcut.Targetpath = target_file_path # The shortcut target file or folder
            shortcut.WorkingDirectory = work_dir # The parent folder of your file
            shortcut.save()
        except Exception as ex:
            print("Error in _create_short_cut" + str(ex))

def _install_pyaudio_windows():
    #Install pyaudio
    try:
        import pyaudio
    except:
        sys_version = str(sys.version[0:6]).strip()
        
        if "3.7" in sys_version :
            cmd = "https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/Wheels/PyAudio-0.2.11-cp37-cp37m-win_amd64.whl?raw=true"
        elif "3.8" in sys_version :
            cmd = "https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/Wheels/PyAudio-0.2.11-cp38-cp38-win_amd64.whl?raw=true"
        elif "3.9" in sys_version :
            cmd = "https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/Wheels/PyAudio-0.2.11-cp39-cp39-win_amd64.whl?raw=true"

        time.sleep(5)

        try:
            os.system("pip install " + cmd)
        except:
            print("Please install appropriate driver from : https://github.com/ClointFusion/Image_ICONS_GIFs/tree/main/Wheels")

        import pyaudio

#Windows OS specific packages
def _load_missing_python_packages_windows():
    
    """
    Installs Windows OS specific python packages
    """       
    list_of_required_packages = ["pywin32","PyGetWindow","pywinauto","comtypes","xlwings","win10toast-click","winshell"] 
    try:
        reqs = subprocess.check_output([sys.executable, '-m', 'pip', 'list'])
        installed_packages = [r.decode().split('==')[0] for r in reqs.split()]
        missing_packages = ' '.join(list(set(list_of_required_packages)-set(installed_packages)))
        
        if missing_packages:
            print("{} package(s) are missing".format(missing_packages)) 
            
            if "comtypes" in missing_packages:
                os.system("{} -m pip install comtypes==1.1.7".format(sys.executable))
            else:
                os.system("{} -m pip install --upgrade pip".format(sys.executable))
            
            cmd = "pip install --upgrade {}".format(missing_packages)
            
            os.system(cmd) 

    except Exception as ex:
        print("Error in _load_missing_python_packages_windows="+str(ex))

def _getCurrentVersion():
    global c_version
    try:
        if os_name == windows_os:
            c_version = os.popen('pip show ClointFusion | findstr "Version"').read()
        elif os_name == linux_os:
            c_version = os.popen('pip3 show ClointFusion | grep "Version"').read()

        if c_version:
            c_version = str(c_version).split(":")[1].strip()
        
    except Exception as ex:
        print("Error in _getCurrentVersion = " + str(ex))

    return c_version

def _getServerVersion():
    global s_version
    try:
        response = requests.get(f'https://pypi.org/pypi/ClointFusion/json')
        s_version = response.json()['info']['version']
    except Warning:
        pass
    except Exception as ex:
        print("Error in _getServerVersion = " + str(ex))

    return s_version

#Linux OS specific packages
def _load_missing_python_packages_linux():
    """
    Installs Linux OS specific python packages
    """       
    list_of_required_packages = ["comtypes"]
    
    additional_ubuntu_packages = "sudo apt-get install python3-tk python3-dev fonts-symbola scrot libcairo2-dev libjpeg-dev libgif-dev libgirepository1.0-dev python3-apt python3-xlib espeak ffmpeg libespeak1 python-pyaudio python3-pyaudio xsel"
    try:
        os.system(additional_ubuntu_packages)
        os.system("xhost +SI:localuser:root")
        os.system(f"xhost +SI:localuser:{str(os.getlogin())}")
        reqs = subprocess.check_output([sys.executable, '-m', 'pip', 'list'])
        installed_packages = [r.decode().split('==')[0] for r in reqs.split()]
        missing_packages = ' '.join(list(set(list_of_required_packages)-set(installed_packages)))
        if missing_packages:
            print("{} package(s) are missing".format(missing_packages)) 
            
            if "comtypes" in missing_packages:
                os.system("sudo {} -m pip install comtypes==1.1.7".format(sys.executable))
            else:
                os.system("sudo {} -m pip install --upgrade pip".format(sys.executable))
            
            cmd = "sudo pip3 install --upgrade {}".format(missing_packages)
            
            os.system(cmd) 

    except Exception as ex:
        print("Error in _load_missing_python_packages_linux="+str(ex))

def _download_cloint_ico_png():
    """
    Internal function to download ClointFusion ICON from GitHub
    """
    global clointfusion_directory, cf_icon_file_path, cf_icon_cdt_file_path, cf_logo_file_path, cf_splash_png_path
    try:
        folder_create(os.path.join(clointfusion_directory,"Logo_Icons")) 

    except Exception as ex: #Ask ADMIN Rights if REQUIRED
        print("Error in _download_cloint_ico_png="+str(ex))

    try:
        if not os.path.exists(str(cf_icon_file_path)):
            urllib.request.urlretrieve('https://raw.githubusercontent.com/ClointFusion/Image_ICONS_GIFs/main/Cloint-ICON.ico',str(cf_icon_file_path))

        if not os.path.exists(str(cf_icon_cdt_file_path)):
            urllib.request.urlretrieve('https://raw.githubusercontent.com/ClointFusion/Image_ICONS_GIFs/main/Cloint-ICON-CDT.ico',str(cf_icon_cdt_file_path))

        if not os.path.exists(str(cf_logo_file_path)):
            urllib.request.urlretrieve('https://raw.githubusercontent.com/ClointFusion/Image_ICONS_GIFs/main/Cloint-LOGO.PNG',str(cf_logo_file_path))

        if not os.path.exists(str(cf_splash_png_path)):
            urllib.request.urlretrieve('https://raw.githubusercontent.com/ClointFusion/Image_ICONS_GIFs/main/Splash.png',str(cf_splash_png_path))

        #BOL related Audio files
        if not os.path.exists(str(Path(os.path.join(clointfusion_directory,"Logo_Icons","Applause.wav")))):
            urllib.request.urlretrieve('https://raw.githubusercontent.com/ClointFusion/Image_ICONS_GIFs/main/applause.wav',(str(Path(os.path.join(clointfusion_directory,"Logo_Icons","Applause.wav")))))

    except Exception as ex:
        print("Error while downloading Cloint ICOn/LOGO = "+str(ex))

# API Functions
def gf(os_hn_ip, time_taken, file_contents):
    try:
        URL = 'https://docs.google.com/forms/d/e/1FAIpQLSehRuz_RWJDcqZMAWRPMOfV7CVZB7PjFruXZtQKXO1Q81jOgw/formResponse?usp=pp_url&entry.1012698071={}&entry.2046783065={}&entry.705740227={}&submit=Submit'.format(str(system_uuid), os_hn_ip + ";" + str(time_taken),file_contents)
        driver = browser_activate(URL)
        return driver
    except:
        pass

def sso():
    try:
        URL = "https://api.clointfusion.com/cf/google/login_process?uuid={}".format(str(system_uuid))
        
        if os_name == linux_os and is_root():
            os.system("xhost +SI:localuser:root")
            subprocess.call('sudo pkill -9 chrome', shell=True)
            os.system(f'google-chrome {URL} --profile-directory="Default" --no-sandbox')
        else:
            webbrowser.open_new(URL)
    except:
        pass

def vst():
    try:
        return requests.post(
            'https://api.clointfusion.com/verify_self_test',
            data={'uuid': str(system_uuid)
                  },
        )
    except:
        pass

def sfn():
    try:
        return requests.post(
            'https://api.clointfusion.com/update_last_month',
            data={
                'last_self_test_month': "0", 
                'uuid': str(system_uuid)
                },
        )
    except:
        pass

def ast():
    try:
        if _self_test(): # Sending email after self test.
            return requests.post(
                'https://api.clointfusion.com/update_last_month',
                data={
                    'last_self_test_month': "1",
                    'uuid': str(system_uuid),
                    },
            )
    except:
        pass

def broadcast_message():
    return requests.post('https://api.clointfusion.com/broadcast_msg',data={'uuid': str(system_uuid)},)

def auto_liker():
    return requests.post('https://api.clointfusion.com/auto_liker',data={'uuid': str(system_uuid)},)

def crash_report(ex):
    """
    Encrypted function used in try/catch block of every function, across all functions in all .py files
    Captures error stack and emails Code Maintainer along with python code for analysis
    """
    send_email_url = "https://api.clointfusion.com/send_gmail"
    import requests
    import pandas as pd
    try:
        df = pd.read_sql('Select * from SYS_CONFIG',connct)
        
        MSG = "<br>Machine Specs : " + str(df.to_dict('records')) + "<br><br>" + "Error Details : " + str(ex)
        # print(MSG)
        resp = requests.post(send_email_url, data={'to_addrs':'mmv.clointfusion@gmail.com','subject':'CCE | Exception | {}'.format(str(os_name.upper())),'message': MSG})

    except Exception as ex:
        print("Error in CrashReport="+str(ex))

# crash_report(traceback.format_exception(*sys.exc_info(),limit=None, chain=True))

# Utility Functions
class DisableLogger():
    def __enter__(self):
       logging.disable(logging.CRITICAL)
    def __exit__(self, exit_type, exit_value, exit_traceback):
       logging.disable(logging.NOTSET)

def folder_create(strFolderPath=""):
    """
    while making leaf directory if any intermediate-level directory is missing,
    folder_create() method will create them all.

    Parameters:
        folderPath (str) : path to the folder where the folder is to be created.

    For example consider the following path:
    
    """
    try:
        if not os.path.exists(strFolderPath):
            os.makedirs(strFolderPath, exist_ok=True)

    except Exception as ex:
        print("Error in folder_create="+str(ex))

def is_root():
    if os_name == linux_os:
        return os.geteuid() == 0

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
    global browser_driver, helium_service_launched

    try:
        # To clear previous instances of chrome
        if clear_previous_instances:
            if os_name == windows_os:
                try:
                    subprocess.call('TASKKILL /IM chrome.exe', stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)
                except Exception as ex:
                    print(f"Error while closing previous chrome instances. {ex}")
            elif os_name == mac_os:
                try:
                    subprocess.call('pkill "Google Chrome"', shell=True)
                except Exception as ex:
                    print(f"Error while closing previous chrome instances. {ex}")
            elif os_name == linux_os:
                try:
                    subprocess.call('sudo pkill -9 chrome', shell=True)
                except Exception as ex:
                    print(f"Error while closing previous chrome instances. {ex}")

        options = Options()
        options.add_argument("--start-maximized")
        options.add_experimental_option('excludeSwitches', ['enable-logging','enable-automation'])
        if os_name == linux_os:
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage') 
        if incognito:
            options.add_argument("--incognito")
        if not dummy_browser:
            if os_name == windows_os:
                options.add_argument("user-data-dir=C:\\Users\\{}\\AppData\\Local\\Google\\Chrome\\User Data".format(os.getlogin()))
            elif os_name == mac_os:
                options.add_argument("user-data-dir=/Users/{}/Library/Application/Support/Google/Chrome/User Data".format(os.getlogin()))
            options.add_argument(f"profile-directory={profile}")
        #  Set the download path
        if files_download_path != '':
            prefs = {
                'download.default_directory': files_download_path,
                "download.prompt_for_download": False,
                'download.directory_upgrade': True,
                "safebrowsing.enabled": False
            }
            options.add_experimental_option('prefs', prefs)

        try:
            with DisableLogger():
                browser_driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
            browser.set_driver(browser_driver)
            if url:
                browser.go_to(url)
            if not url:
                browser.go_to("https://sites.google.com/view/clointfusion-hackathon")
            browser.Config.implicit_wait_secs = 120
        except Exception as ex:
            print(f"Error while browser_activate: {str(ex)}")
    except Exception as ex:
        print("Error in launch_website_h = " + str(ex))
        browser.kill_browser()
    finally:
        return browser_driver

def database_connect(db_name):
    try:
        conn = sqlite3.connect(db_name)
        c = conn.cursor()
        return conn, c
    except Exception as ex:
        print(ex)

def get_details():
    global cursr, c_version, s_version, username, email
    try:
        username = ""
        email = ""
        data = cursr.execute("SELECT username, email from CF_VALUES")
        for row in data:
            username =  row[0]
            email =  row[1]
        if username == "username" or email == "email":
            _, username, email = str(vst().text).split('#')
        return username , c_version, s_version, email
    
    except Exception as ex:
        print("Error in get_details = ", str(ex))

def get_uuid():
    try:
        return system_uuid
    except:
        pass

def _update_version(c_version,s_version):
    if c_version:
        print('You are using version {}, however version {} is available !'.format(c_version,s_version))
    else:
        print(f'Version {s_version} found on PyPi Server')

    print('\nUpgrading to latest version...Please wait a moment...\n')
    try:
        script =r"""
import os, platform, time, sys
exe = str(sys.executable)
os_name = str(platform.system()).lower()
windows_os = "windows"
linux_os = "linux"
mac_os = "darwin"

print('Please wait while i update the ClointFusion for you.')
time.sleep(5)
try:
    if os_name == windows_os:
        os.system(f"{exe} -m pip install --upgrade pip")
        os.system(f"pip install -U ClointFusion --no-warn-script-location")
        os.system("cls")
        os.system(f'{exe} -c "import ClointFusion"')

    elif os_name == linux_os:
        os.system("sudo python3 -m pip install --upgrade pip")
        os.system("sudo pip3 install -U ClointFusion")
        os.system("clear")
        os.system('sudo python3 -c "import ClointFusion"')
except Exception as ex:
    print("Error in _update_version = " + str(ex))

        """
        with open(update_path, 'w') as f:
            f.writelines(script)
        print("Starting update...")
        time.sleep(5)
        
        try:
            os.startfile(update_path)
            
        except Exception as ex:
            print(ex, "update_path")
    except Exception as ex:
        print("Please Upgrade ClointFusion using: pip install -U ClointFusion", str(ex))

def verify_version(c_version, s_version):
    try:
        global cursr, connct, username, email
        if c_version < s_version:
            sfn()
            cursr.execute("UPDATE CF_VALUES set FIRST_RUN = 'True', SELF_TEST = 'True', UPDATING = 'True'   where ID = 1")
            connct.commit()
            _update_version(c_version, s_version)
        else:
            resp = vst()
            last_updated_on_month, username, email = str(resp.text).split('#')
            cursr.execute("UPDATE CF_VALUES SET USERNAME=?, EMAIL=? WHERE id = 1", (username, email))
            connct.commit()
            if last_updated_on_month != "0":
                cursr.execute("UPDATE CF_VALUES set SELF_TEST = 'False', UPDATING = 'False' where ID = 1")
                connct.commit()
                
    except Exception as ex:
        print("Error in verify_version = " + str(ex))
        
get_current_version_thread = threading.Thread(target=_getCurrentVersion, name="GetCurrentVersion")
get_current_version_thread.start()

get_server_version_thread = threading.Thread(target=_getServerVersion, name="GetServerVersion")
get_server_version_thread.start()

get_current_version_thread.join()
get_server_version_thread.join()

_download_cloint_ico_png()

folder_create(clointfusion_directory)
folder_create(config_folder_path)
folder_create(log_path)
folder_create(img_folder_path)
folder_create(batch_file_path)
folder_create(config_folder_path)
folder_create(error_screen_shots_path)
folder_create(output_folder_path)

try:
    connct, cursr = database_connect(db_file_path)
except Exception as ex:
    print("Error in connecting database = ", str(ex))

try:  
    cursr.execute('''CREATE TABLE CF_VALUES
        (ID INT PRIMARY KEY     NOT NULL,
        FIRST_RUN           TEXT    NOT NULL,
        BOL            INT     NOT NULL,
        SELF_TEST        TEXT,
        USERNAME        TEXT,
        EMAIL        TEXT,
        UPDATING        TEXT);''')
    cursr.execute("INSERT INTO CF_VALUES (ID,FIRST_RUN,BOL,SELF_TEST, USERNAME, EMAIL, UPDATING) \
    VALUES (1, 'True', 1, 'True', 'username', 'email', 'True');")
    connct.commit()
except sqlite3.OperationalError:
    pass
except Exception as ex :
    print(f"Exception: {ex}")

try:
    if os.path.exists(update_path):
        os.remove(update_path)
except Exception as ex:
    print("Error in removing update file = ", str(ex))

site_pkg_path = _get_site_packages_path()
verify_version(c_version, s_version)

data = cursr.execute("SELECT first_run,self_test from CF_VALUES")
for row in data:
    FIRST_RUN =  row[0]

if FIRST_RUN == "True":
    
    try:
        if os_name == windows_os:

            _load_missing_python_packages_windows()
            _install_pyaudio_windows()

            bre_file_path = f"{site_pkg_path}" + '\ClointFusion\BRE_WHM.pyw'
            notifications_path = f"{site_pkg_path}" + '\ClointFusion\cf_notification.pyw'
            cf_folder = f"{site_pkg_path}" + '\ClointFusion'

            current_user = str(str(Path.home()).split("\\")[2])

            short_cut_path_bre = r"C:\Users\{}\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup".format(current_user) + "\CF_Tray.lnk"
            short_cut_path_noti = r"C:\Users\{}\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup".format(current_user) + "\CF_Notification.lnk"

            try:
                _update_registry(short_cut_path_bre)
                _update_registry(short_cut_path_noti)
            except Exception as ex:
                print("Error in updating registry = ", str(ex))

            _create_short_cut(short_cut_path_bre,bre_file_path,cf_folder)
            _create_short_cut(short_cut_path_noti,notifications_path,cf_folder)
            cursr.execute("UPDATE CF_VALUES set FIRST_RUN = 'False' where ID = 1")
            connct.commit()
        elif os_name == linux_os:
            _load_missing_python_packages_linux()
        else:
            pass
    except Exception as ex:
        print("Error in FIRST_RUN setup ="+str(ex))

# Welcome to ClointFusion
welcome_msg = "Welcome to ClointFusion."

cf_defination = ["ClointFusion is a python based RPA tool",
                 "ClointFusion is a Automation tool for Common Man.",
                 "ClointFusion is an Automation Framework.",
                 "ClointFusion helps Common Man to automate their boring tasks.",]

cf_defination_msg = random.choice(cf_defination)