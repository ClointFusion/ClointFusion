import webbrowser
from win10toast_click import ToastNotifier 
import requests
import schedule, platform, os, time
from pathlib import Path
from datetime import datetime, timedelta
from ClointFusion import selft

toaster = ToastNotifier()

os_name = str(platform.system()).lower()
windows_os = "windows"
linux_os = "linux"
mac_os = "darwin"

if os_name == windows_os:
    clointfusion_directory = r"C:\Users\{}\ClointFusion".format(str(os.getlogin()))
elif os_name == linux_os:
    clointfusion_directory = r"/home/{}/ClointFusion".format(str(os.getlogin()))
elif os_name == mac_os:
    clointfusion_directory = r"/Users/{}/ClointFusion".format(str(os.getlogin()))

img_folder_path =  Path(os.path.join(clointfusion_directory, "Images"))
config_folder_path = Path(os.path.join(clointfusion_directory, "Config_Files"))
cf_splash_png_path = Path(os.path.join(clointfusion_directory,"Logo_Icons","Splash.PNG"))
cf_icon_cdt_file_path = os.path.join(clointfusion_directory,"Logo_Icons","Cloint-ICON-CDT.ico")

def _getServerVersion():
    global s_version
    try:
        response = requests.get(f'https://pypi.org/pypi/ClointFusion/json')
        s_version = response.json()['info']['version']
    except Warning:
        pass

    return s_version

def _getCurrentVersion():
    global c_version
    try:
        if os_name == windows_os:
            c_version = os.popen('pip show ClointFusion | findstr "Version"').read()
        elif os_name == linux_os:
            c_version = os.popen('pip3 show ClointFusion | grep "Version"').read()

        c_version = str(c_version).split(":")[1].strip()
    except:
        pass

    return c_version

def show_toast_notification_if_new_version_is_available():
    if os_name == windows_os:
        s_version = _getServerVersion()
        c_version = _getCurrentVersion()
        if c_version < s_version:
            toaster.show_toast(
                "ClointFusion", 
                "New version {} is available now ! Click to update".format(s_version), 
                icon_path=cf_icon_cdt_file_path,
                duration=None,
                threaded=True, 
                callback_on_click=lambda: os.system('cf') # click notification to run function 
            )
        else:
            try:
                resp = selft.broadcast_message()
                name, _, _ = selft.get_details()
                resp = eval(resp.text)
                server_msg = resp['msg']
                server_url = resp['url']
                server_date, server_month = resp['dt'].split('/')[0], resp['dt'].split('/')[1]
                today_date, today_month  = datetime.now().day, datetime.now().month
                
                if today_date <= server_date or today_month < server_month:
                    toaster.show_toast(
                        "ClointFusion", 
                        f"Hai {name}, \n{server_msg}\nClick for more detail.",
                        icon_path=cf_icon_cdt_file_path,
                        duration=None,
                        threaded=True,
                        callback_on_click=lambda: webbrowser.open(server_url)
                    )
            except:
                pass
            
schedule.every(5).hour.do(show_toast_notification_if_new_version_is_available)

# Server Broadcast
while True:
    schedule.run_pending()
    time.sleep(60) #Check Every minute
