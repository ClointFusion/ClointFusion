from win10toast_click import ToastNotifier
import requests
import schedule
from ClointFusion import selft
import datetime
from datetime import timedelta
import webbrowser,platform, os, time, datetime
import pyinspect as pi
from rich import pretty
from pathlib import Path
import sys


windows_os = "windows"
linux_os = "linux"
mac_os = "darwin"
os_name = str(platform.system()).lower()

if os_name == windows_os:
    clointfusion_directory = r"C:\Users\{}\ClointFusion".format(str(os.getlogin()))
elif os_name == linux_os:
    clointfusion_directory = r"/home/{}/ClointFusion".format(str(os.getlogin()))
elif os_name == mac_os:
    clointfusion_directory = r"/Users/{}/ClointFusion".format(str(os.getlogin()))

cf_icon_cdt_file_path = os.path.join(clointfusion_directory,"Logo_Icons","Cloint-ICON-CDT.ico")
pi.install_traceback(hide_locals=True,relevant_only=True,enable_prompt=True)
pretty.install()

toaster = ToastNotifier()

local_msg = ""
local_url = ""
local_date = datetime.datetime.now().strftime("%d/%m/%Y")
server_date = ""

def act_on_click():
    global local_msg

    if "new version" in str(local_msg).lower():\
        os.system('cf')
    else:
        webbrowser.open(local_url)

    local_msg = "done"

def show_toast_notification_if_new_msg_is_available():
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
                # print(local_msg)
                if server_date < local_date:
                    local_msg = "done"
            
            if local_msg != "done" and local_date == server_date:
                toaster.show_toast(
                    "ClointFusion", 
                    f"Hi {name}, \n{local_msg}\nClick here for more details",
                    icon_path=cf_icon_cdt_file_path,
                    duration=30,
                    threaded=True,
                    callback_on_click=lambda: act_on_click()
                )
            
            if local_date > server_date:
                local_msg = ""

        except Exception as ex:
            print("Error in show_toast_notification_if_new_msg_is_available" + str(ex))

# # # # Actual Values # # # #
# schedule.every(1).to(6).hours.do(show_toast_notification_if_new_msg_is_available)
# Server Broadcast
# while True:
#     schedule.run_pending()
#     time.sleep(60) #Check Every minute

# # # # For test uncomment below # # # #
# schedule.every(5).to(15).seconds.do(show_toast_notification_if_new_msg_is_available)
# # Server Broadcast
# i = 1
# while True:
    
#     print("run", i)
#     i += 1
#     schedule.run_pending()
#     time.sleep(2) #Check Every minute

show_toast_notification_if_new_msg_is_available()
