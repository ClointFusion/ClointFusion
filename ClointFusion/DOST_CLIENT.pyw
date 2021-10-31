from selenium.common.exceptions import TimeoutException, WebDriverException
import logging
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import helium as browser
from rich.text import Text
from rich import print
from rich.console import Console
from ClointFusion import selft
import requests, os, subprocess, platform, time, sys, traceback
from rich import pretty
import pyinspect as pi

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

pi.install_traceback(hide_locals=True,relevant_only=True,enable_prompt=True)
pretty.install()

console = Console()

start = True
found = False
script = True

website = "https://dost.clointfusion.com"
temp_code = clointfusion_directory + "\Config_Files\dost_code.py"

class DisableLogger():
    def __enter__(self):
       logging.disable(logging.CRITICAL)
    def __exit__(self, exit_type, exit_value, exit_traceback):
       logging.disable(logging.NOTSET)

def browser_activate(url="", files_download_path='', dummy_browser=True, incognito=False,
                     clear_previous_instances=False, profile="Default"):
    """ This function is used to activate the browser.  """
    try:
    # To clear previous instances of chrome
        if clear_previous_instances:
            if os_name == windows_os:
                try:
                    subprocess.call('TASKKILL /IM chrome.exe', stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)
                except Exception as ex:
                    selft.crash_report(traceback.format_exception(*sys.exc_info(),limit=None, chain=True))
                    print(f"Error while closing previous chrome instances. {ex}")
            elif os_name == mac_os:
                try:
                    subprocess.call('pkill "Google Chrome"', shell=True)
                except Exception as ex:
                    selft.crash_report(traceback.format_exception(*sys.exc_info(),limit=None, chain=True))
                    print(f"Error while closing previous chrome instances. {ex}")
            elif os_name == linux_os:
                try:
                    subprocess.call('sudo pkill -9 chrome', shell=True)
                except Exception as ex:
                    selft.crash_report(traceback.format_exception(*sys.exc_info(),limit=None, chain=True))
                    print(f"Error while closing previous chrome instances. {ex}")
            time.sleep(10)

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
            browser.Config.implicit_wait_secs = 20
        except Exception as ex:
            selft.crash_report(traceback.format_exception(*sys.exc_info(),limit=None, chain=True))
            print(f"Error while browser_activate: {str(ex)}")
    except Exception as ex:
        selft.crash_report(traceback.format_exception(*sys.exc_info(),limit=None, chain=True))
        print("Error in launch_website_h = " + str(ex))
        browser.kill_browser()
    finally:
        return browser_driver

def clear_screen():
    """
    Clears Python Interpreter Terminal Window Screen
    """
    try:
        command = 'cls' if os.name in ('nt', 'dos') else 'clear'
        os.system(command)
    except:
        pass

def run_program(website):
  code = requests.get(f"{website}/cf_id_get/{uuid}").json()["code"]
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

try:
    uuid = selft.get_uuid()
    text = Text("Welcome to DOST,")
    text.stylize("bold magenta")
    text.append(" you drag and drop, we do the rest. Happy Automation!")
    console.print(text)

    web_driver = browser_activate(url=f"{website}/cf_id/{uuid}", dummy_browser=False, clear_previous_instances=True)
    browser.set_driver(web_driver)
    browser.Config.implicit_wait_secs = 5
except Exception as ex:
    selft.crash_report(traceback.format_exception(*sys.exc_info(),limit=None, chain=True))
    print(f"Error in UUID: {str(ex)}")

try:
    with console.status("DOST client running...\n") as status:
        while start:
            try:
                if not found:
                    run_btn = browser.find_all(browser.S('//*[@id="cf_run"]'))
                    if run_btn:
                        if script:
                            web_driver.execute_script('localStorage.setItem("client", true)')
                            script = False
                        found = run_btn[0]
                if found:
                    browser.wait_until(browser.Text("Running Program..").exists)
                    if browser.Text("Running Program...").exists:
                        browser.wait_until(lambda: not browser.Text("Running Program..").exists())
                        status.update("Running your bot...\n")
                        while run_program(website):
                            continue
                        status.update("DOST client running...\n")
                        found = False
            except TimeoutException:
                found = False
            except WebDriverException:
                browser.kill_browser()
                start = False
            except IndexError:
                browser.kill_browser()
                start = False
            except Exception as ex:
                selft.crash_report(traceback.format_exception(*sys.exc_info(),limit=None, chain=True))
                print("Error in DOST_Client.pyw="+str(ex))

except Exception as ex:
    selft.crash_report(traceback.format_exception(*sys.exc_info(),limit=None, chain=True))
    print(f"Error in DOST_Client: {str(ex)}")
