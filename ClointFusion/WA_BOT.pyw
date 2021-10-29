import ClointFusion as cf
import traceback
from urllib.parse import quote
import pyinspect as pi
from rich import pretty

pi.install_traceback(hide_locals=True,relevant_only=True,enable_prompt=True)
pretty.install()

def send_wa_msg(mobile_number,name, msg):
    try:
        print("Sending WA MSG to ", mobile_number, name, msg)
        url_qry = "https://web.whatsapp.com/send?phone=" + str(mobile_number) + "&text=" + quote(msg)
        cf.browser_navigate_h(url=url_qry)
        cf.browser_mouse_click_h(msg)
        cf.time.sleep(2)
        cf.browser_hit_enter_h()
        cf.time.sleep(5)
        print("WA MSG sent.")
    except Exception as ex:
        cf.selft.crash_report(traceback.format_exception(*cf.sys.exc_info(),limit=None, chain=True))
        print("Error while send_wa_msg "+ str(ex))

def send_wa_in_loop(excel_path):
    try:
        df = cf.pd.read_excel(excel_path,engine='openpyxl')
        for _, row in df.iterrows():
            mobile_number = row[0]
            name = row[1]
            msg = row[2]

            if "+91" not in str(mobile_number) and len(str(mobile_number)) == 10:
                mobile_number = "91" + str(mobile_number)

            elif "+" not in str(mobile_number) and len(str(mobile_number)) == 12:
                mobile_number = str(mobile_number)
                
            elif "+" in str(mobile_number) and len(str(mobile_number)) == 13:
                mobile_number = str(mobile_number).strip("+")

            try:
                send_wa_msg(mobile_number,name, msg)
            except Exception as ex:
                cf.selft.crash_report(traceback.format_exception(*sys.exc_info(),limit=None, chain=True))
                print("Error in send_wa_in_loop", str(ex))
        cf.browser_quit_h()
    except Exception as ex:
        cf.selft.crash_report(traceback.format_exception(*sys.exc_info(),limit=None, chain=True))
        print("Error while send_wa_in_loop "+ str(ex))
        # browser.Alert().dismiss()

def shoot_msg(excel_path):        
    try:
        cf.browser_activate("web.whatsapp.com", dummy_browser=False, clear_previous_instances=True)
        cf.browser_set_waiting_time(30)

        try:
            logined = True if str(cf.browser_locate_element_h('//*[@id="app"]/div[1]/div[1]/div[4]/div/div/div[2]/div[2]/div[2]/div/a', get_text=True)).lower() == "get it here" else False

            if not logined:
                logined = True if str(cf.browser_locate_element_h('//*[@id="app"]/div[1]/div[1]/div[4]/div/div/div[2]/div[3]/div[2]/div/a', get_text=True)).lower() == "get it here" else False

        except Exception as ex:
            cf.selft.crash_report(traceback.format_exception(*sys.exc_info(),limit=None, chain=True))
            print("User is not logged in.")
        

        if not logined:    
            cf.text_to_speech("Please scan the QR code and login to whatsapp web.", show=False)
            while not logined:
                try:
                    logined = True if str(cf.browser_locate_element_h('//*[@id="app"]/div[1]/div[1]/div[4]/div/div/div[2]/div[2]/div[2]/div/a', get_text=True)).lower() == "get it here" else False
                    if not logined:
                        logined = True if str(cf.browser_locate_element_h('//*[@id="app"]/div[1]/div[1]/div[4]/div/div/div[2]/div[3]/div[2]/div/a', get_text=True)).lower() == "get it here" else False
                except Exception as ex:
                    cf.selft.crash_report(traceback.format_exception(*sys.exc_info(),limit=None, chain=True))
                    print("Waiting for you to log in.")
        
        cf.text_to_speech("OK, Let me send the messages", show=False)
        cf.browser_set_waiting_time(60)
        
        send_wa_in_loop(excel_path)
    except Exception as ex:
        cf.selft.crash_report(traceback.format_exception(*sys.exc_info(),limit=None, chain=True))
        print("Error while shoot_msg "+ str(ex))

if len(cf.sys.argv) > 1:
    shoot_msg(cf.sys.argv[1])