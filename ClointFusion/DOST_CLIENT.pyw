from selenium.common.exceptions import InvalidArgumentException, TimeoutException, WebDriverException
import logging
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import helium as browser
from rich.text import Text
from rich import print
from rich.console import Console
try:
    from ClointFusion import selft
except:
    import selft
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

def _get_site_packages_path():
    """
    Returns Site-Packages Path
    """
    try:
        import site
        if os_name == windows_os:
            site_packages_path = next(p for p in site.getsitepackages() if 'site-packages' in p)
        else:
            site_packages_path = site.getsitepackages()[0]
    except:
        site_packages_path = site.getsitepackages()[0]
    return str(site_packages_path)

def _welcome_to_clointfusion():
    global user_name
    """
    Internal Function to display welcome message & push a notification to ClointFusion Slack
    """
    from pyfiglet import Figlet
    version = "(Version: 1.1.1)"
    import random


    messages_list = ['Where would I be without a friend like you?', 'I appreciate what you did.', 'Thank you for thinking of me.', 'Thank you for your time today.', 'I am so thankful for what you did here', 'I really appreciate your help. Thank you.', "You know, if you're reading this, you're in the top 1% of smart people.", 'We know the world is full of choices. Yet you picked us, Thank you very much.', 'Thank you. We hope your experience was excellent and we can’t wait to see you again soon.', 'We hope you are happy with our tool, if not we are just an e-mail away at clointfusion@cloint.com. We will be pleased to hear from you.', 'ClointFusion would like to thank excellent users like you for your support. We couldn’t do it without you!', 'Thank you for your business, your trust, and your confidence. It is our pleasure to work with you.', 'We take pride in your business with us. Thank you!', 'It has been our pleasure to serve you, and we hope we see you again soon.', 'We value your trust and confidence in us and sincerely appreciate you!', 'Your satisfaction is our greatest concern!', 'Your confidence in us is greatly appreciated!', 'We are excited to serve you first!', 'Thank you for keeping us informed about how best to serve your needs. Together, we can make this history.', 'Our brand innovation wouldn’t have been possible if you didn’t give us feedback about our services.', 'Thank you so much for playing a pivotal role in our growth. We’ll make sure we continue to put your needs first as our company expands and improves.', 'We are exceedingly pleased to find people we can always count on. Thank you for being one of our loyal and trusted clients.', ]
    message = random.choice(messages_list)
    
    print()
    f = Figlet(font='small', width=150)
    console.print(f.renderText("ClointFusion Community Edition"))
    print(message + "\n")

site_packages_path = _get_site_packages_path()

python_version = str(sys.version_info.major)
class DisableLogger():
    def __enter__(self):
       logging.disable(logging.CRITICAL)
    def __exit__(self, exit_type, exit_value, exit_traceback):
       logging.disable(logging.NOTSET)

def browser_activate(url="", files_download_path='', dummy_browser=True,
                     clear_previous_instances=False, profile="Default"):
    """ This function is used to activate the browser.  """
    try:
        browser_driver = ""
    # To clear previous instances of chrome
        if clear_previous_instances:
            try:
                subprocess.call('TASKKILL /IM chrome.exe', stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)
            except Exception as ex:
                selft.crash_report(traceback.format_exception(*sys.exc_info(),limit=None, chain=True))
                print(f"Error while closing previous chrome instances. {ex}")
            time.sleep(15)

        options = Options()
        options.add_argument("--start-maximized")
        options.add_experimental_option('excludeSwitches', ['enable-logging','enable-automation'])
        # options.add_experimental_option("detach", True)

        if not dummy_browser:
            if os_name == windows_os:
                options.add_argument("user-data-dir=C:\\Users\\{}\\AppData\\Local\\Google\\Chrome\\User Data".format(os.getlogin()))
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
            browser.Config.implicit_wait_secs = 5
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
        command = 'cls' if os.name in ('nt', 'dos') else 'clear'
        os.system(command)
    except:
        pass

def run_program(website):
    code = requests.get(f"{website}/cf_id_get/{uuid}").json()["code"]
    
    try:
        
        f_code = "import pyinspect as pi\nimport sys, traceback\nfrom rich import pretty\nfrom ClointFusion.ClointFusion import selft\npi.install_traceback(hide_locals=True,relevant_only=True,enable_prompt=True)\npretty.install()\ntry:\n"
        for line in code.split("\n"):
            f_code += "    " + line + "\n"
        f_code += "except Exception as ex:\n    selft.crash_report(traceback.format_exception(*sys.exc_info(),limit=None, chain=True))\n    print(f'Error in Running DOST Program: {str(ex)}')"

        with open(temp_code, 'w') as fp:
            fp.write(f_code + "\n" + r"print('\n')")
        status = exe_code(temp_code)
    
    # code = requests.get(f"{website}/cf_id_get/{uuid}").json()["code"]
    # try:
    #     with open(temp_code, 'w') as fp:
    #         fp.write(code + "\n" + r"print('\n')")
    #     status = exe_code(temp_code)
    except:
        pass
    return status

def exe_code(path):
    clear_screen()
    cmd = f'python "{path}"'
    os.system(cmd)
    return False

if os_name == windows_os:

    try:
        profile = "Default"
        if len(sys.argv) > 1:
            profile = "".join([character for character in sys.argv[1] if character.isalnum() or character == ' '])
        clear_screen()
        _welcome_to_clointfusion()
        with console.status(f"Launching DOST client with Profile : {profile} .\n") as status:
            uuid = selft.get_uuid()
            text = Text("Welcome to DOST,")
            text.stylize("bold magenta")
            text.append(" you drag and drop, we do the rest. Happy Automation!\n\n")
            console.print(text)
            web_driver = browser_activate(url=f"{website}/cf_id/{uuid}", dummy_browser=False, profile=profile, clear_previous_instances=True)
            browser.set_driver(web_driver)
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
                        browser.wait_until(browser.Text("Running BOT in your terminal, Please wait..").exists)
                        if browser.Text("Running BOT in your terminal, Please wait..").exists:
                            browser.wait_until(lambda: not browser.Text("Running BOT in your terminal, Please wait..").exists())
                            
                            status.update("Running your bot...")
                            while run_program(website):
                                continue
                            status.update("DOST client running...")
                            found = False
                except TimeoutException:
                    found = False
                except WebDriverException:
                    browser.kill_browser()
                    break
                except IndexError:
                    browser.kill_browser()
                    break
                except RuntimeError:
                    print("Please close all the Google Chrome browser windows, and try again.")
                    browser.kill_browser()
                    break
                except InvalidArgumentException:
                    status.update("Another Chrome Windows is already in use.")
                    time.sleep(5)
                except AttributeError:
                    clear_screen()
                    status.update("Another Chrome Windows is already in use.")
                    time.sleep(3)
                    status.update("All Google Chrome browser windows, should be before launching DOST.\n")
                    time.sleep(3)
                    status.update("Please wait, while I close them and try again.\n")
                    try:
                        subprocess.call('TASKKILL /F /IM chrome.exe', stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)
                        time.sleep(5)
                        web_driver = browser_activate(url=f"{website}/cf_id/{uuid}", dummy_browser=False, profile=profile, clear_previous_instances=False)
                        browser.set_driver(web_driver)
                        found = False
                        clear_screen()
                        _welcome_to_clointfusion()
                        status.update("DOST client running...")
                    except Exception as ex:
                        browser.kill_browser()
                        print("Please restart the DOST client, after closing all the Google Chrome windows.")
                        break
                except Exception as ex:
                    # selft.crash_report(traceback.format_exception(*sys.exc_info(),limit=None, chain=True))
                    print("Error in DOST_Client.pyw="+str(ex))
                    break
        print("Thank you for utilizing DOST. I hope you have a good time with it.\nDo you have any suggestions ? Love to hear them, please drop a mail at ClointFusion@cloint.com.\n")
    except Exception as ex:
        selft.crash_report(traceback.format_exception(*sys.exc_info(),limit=None, chain=True))
        print(f"Error in DOST_Client: {str(ex)}")
elif os_name == linux_os:
    code = r"""
from selenium.common.exceptions import TimeoutException, WebDriverException
import logging
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import helium as browser
from rich.text import Text
from rich.console import Console
import requests, os, subprocess, time, traceback,  sys
from rich import pretty
import pyinspect as pi

clointfusion_directory = r"/home/{}/ClointFusion".format(str(os.getlogin()))

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

def clear_screen():
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
  cmd = f'sudo python3 "{path}"'
  os.system(cmd)
  return False

def _welcome_to_clointfusion():
    from pyfiglet import Figlet
    version = "(Version: 1.1.1)"
    import random
    

    messages_list = ['Where would I be without a friend like you?', 'I appreciate what you did.', 'Thank you for thinking of me.', 'Thank you for your time today.', 'I am so thankful for what you did here', 'I really appreciate your help. Thank you.', "You know, if you're reading this, you're in the top 1% of smart people.", 'We know the world is full of choices. Yet you picked us, Thank you very much.', 'Thank you. We hope your experience was excellent and we can’t wait to see you again soon.', 'We hope you are happy with our tool, if not we are just an e-mail away at clointfusion@cloint.com. We will be pleased to hear from you.', 'ClointFusion would like to thank excellent users like you for your support. We couldn’t do it without you!', 'Thank you for your business, your trust, and your confidence. It is our pleasure to work with you.', 'We take pride in your business with us. Thank you!', 'It has been our pleasure to serve you, and we hope we see you again soon.', 'We value your trust and confidence in us and sincerely appreciate you!', 'Your satisfaction is our greatest concern!', 'Your confidence in us is greatly appreciated!', 'We are excited to serve you first!', 'Thank you for keeping us informed about how best to serve your needs. Together, we can make this history.', 'Our brand innovation wouldn’t have been possible if you didn’t give us feedback about our services.', 'Thank you so much for playing a pivotal role in our growth. We’ll make sure we continue to put your needs first as our company expands and improves.', 'We are exceedingly pleased to find people we can always count on. Thank you for being one of our loyal and trusted clients.', ]
    message = random.choice(messages_list)
    
    print()
    f = Figlet(font='small', width=150)
    console.print(f.renderText("ClointFusion Community Edition"))
    print(message + "\n")

def browser_linux(profile = "Default"):
    with console.status(f"Launching DOST client with Profile : {profile}...\n"):
        clear_screen()
        _welcome_to_clointfusion()
        browser_driver = ""
        from subprocess import DEVNULL, STDOUT, Popen
        
        Popen([f'google-chrome --profile-directory="{profile}" --remote-debugging-port=9222'], shell=True,
                stdin=None, stdout=DEVNULL, stderr=STDOUT, close_fds=True)
        time.sleep(5)
        
        options = Options()
        options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        
        with DisableLogger():
            browser_driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)

        return browser_driver

try:
    uuid = str(subprocess.check_output('sudo dmidecode -s system-uuid', shell=True),'utf-8').split('\n')[0].strip()
    text = Text("Welcome to DOST,")
    text.stylize("bold magenta")
    text.append(" you drag and drop, we do the rest. Happy Automation!")
    console.print(text)
    profile = "Default"
    print(sys.argv)
    if len(sys.argv) > 1:
        profile = sys.argv[1]
    web_driver = browser_linux(profile)

    browser.set_driver(web_driver)
    browser.go_to(f"{website}/cf_id/{uuid}")
except Exception as ex:
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
                    browser.wait_until(browser.Text("Running BOT in your terminal, Please wait..").exists)
                    if browser.Text("Running BOT in your terminal, Please wait..").exists:
                        browser.wait_until(lambda: not browser.Text("Running BOT in your terminal, Please wait..").exists())
                        status.update("Running your bot...")
                        while run_program(website):
                            continue
                        status.update("DOST client running...")
                        found = False
            except TimeoutException:
                found = False
            except WebDriverException:
                print("Thank you for utilizing DOST. I hope you have a good time with it.")
                browser.kill_browser()
                start = False
                break
            except IndexError:
                browser.kill_browser()
                start = False
                break
            except RuntimeError:
                print("Please close all the Google Chrome browser windows, and try again.")
                browser.kill_browser()
                break
            except Exception as ex:
                print("Error in DOST_Client.pyw="+str(ex))
                break
        print("Thank you for utilizing DOST. I hope you have a good time with it.\nDo you have any suggestions ? Love to hear them, please drop a mail at ClointFusion@cloint.com.\n")
        
except Exception as ex:
    print(f"Error in DOST_Client: {str(ex)}")
    
    """
    import os
    path_py = f"{os.getcwd()}/dost.py"
    with open (path_py, 'w') as rsh:
        rsh.write(code)
    os.system(f"chmod +x {path_py}")
    
    print(f'Please run the following command to start DOST Client: \npython{python_version} dost.py\nWant to use a different profile, use\npython{python_version} dost.py "Person 1"\n')
    

