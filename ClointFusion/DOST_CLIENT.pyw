from selenium.common.exceptions import TimeoutException, WebDriverException
import requests
import os
import subprocess
import platform
import logging
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import helium as browser
from rich.text import Text
from rich import print
from rich.console import Console

console = Console()
os_name = str(platform.system()).lower()
windows_os = "windows"
linux_os = "linux"
mac_os = "darwin"

# UUID KEY
os_name = str(platform.system()).lower()
if os_name == windows_os:
    uuid = str(subprocess.check_output('wmic csproduct get uuid'), 'utf-8').split('\n')[1].strip()
else:
    uuid = str(subprocess.check_output('sudo dmidecode -s system-uuid', shell=True),'utf-8').split('\n')[0].strip()

clointfusion_directory = r"C:\Users\{}\ClointFusion".format(str(os.getlogin()))
temp_code = clointfusion_directory + "\Config_Files\dost_code.py"

class DisableLogger():
    def __enter__(self):
       logging.disable(logging.CRITICAL)
    def __exit__(self, exit_type, exit_value, exit_traceback):
       logging.disable(logging.NOTSET)


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
    status = False
    browser_driver = ''
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
                    subprocess.call('killall chrome', shell=True)
                except Exception as ex:
                    print(f"Error while closing previous chrome instances. {ex}")

        options = Options()
        options.add_argument("--start-maximized")
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
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

def clear_screen():
    """
    Clears Python Interpreter Terminal Window Screen
    """
    try:
      command = 'clear'
      if os.name in ('nt', 'dos'):  # If Machine is running on Windows, use cls
        command = 'cls'
      os.system(command)
    except:
        pass


web_driver = browser_activate(url=f"dost.clointfusion.com/cf_id/{uuid}", dummy_browser=False, clear_previous_instances=True)
browser.set_driver(web_driver)

found = False

def run_program():
  code = requests.get(f"https://dost.clointfusion.com/cf_id_get/{uuid}").json()["code"]
  try:
    with open(temp_code, 'w') as fp:
        fp.write(code + "\n" + r"print('\n')")
    status = exe_code(temp_code)
  except:
    pass
  return status

def exe_code(path):
  clear_screen()
  cmd = f'python "{path}"'
  os.system(cmd)
  return False


text = Text("Welcome to DOST,")
text.stylize("bold magenta")
text.append(" you drag and drop, we do the rest. Happy Automation")
console.print(text)

with console.status("DOST client running...\n") as status:
  while (True):
    if not found:
      try:
        run_btn = browser.find_all(browser.S('//*[@id="cf_run"]'))
        if run_btn:
          found = run_btn[0]
      except WebDriverException:
        break

    if found:
      try:
        browser.wait_until(browser.Text("Running Program..").exists)
        if browser.Text("Running Program..").exists:
          browser.wait_until(lambda: not browser.Text("Running Program..").exists())
          while run_program():
            continue
          found = False
      except TimeoutException:
        found = False
      except WebDriverException:
        break
  
  
