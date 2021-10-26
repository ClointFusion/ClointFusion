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

local_msg = ""
local_url = ""
local_date = datetime.now().strftime("%d/%m/%Y")
server_date = ""

def act_on_click():
    global local_msg

    if "new version" in str(local_msg).lower():\
        os.system('cf')
    else:
        webbrowser.open(local_url)

    local_msg = "done"

def show_toast_notification_if_new_version_is_available():
    global local_msg, local_url, server_date
    name, _, _, _ = selft.get_details()
    
    if os_name == windows_os:
        try:
            if local_msg == "" : # or server_date == "":
                resp = selft.broadcast_message()
                resp = eval(resp.text)
                local_msg = resp['msg']
                local_url = resp['url']
                # msg_date_n, msg_month_n = int(resp['dt'].split('/')[0]), int(resp['dt'].split('/')[1])
                server_date=resp['dt']

                if server_date < local_date:
                    local_msg = "done"
            
            if local_msg != "done" and local_date == server_date:
                toaster.show_toast(
                    "ClointFusion", 
                    f"Hi {name}, \n{local_msg}\nClick here for more details",
                    icon_path=cf_icon_cdt_file_path,
                    duration=None,
                    threaded=True,
                    callback_on_click=lambda: act_on_click()
                )
        except Exception as ex:
            print("Error in show_toast_notification_if_new_version_is_available" + str(ex))
    
schedule.every(1).to(6).hours.do(show_toast_notification_if_new_version_is_available)

# Server Broadcast
while True:
    schedule.run_pending()
    time.sleep(60) #Check Every minute
