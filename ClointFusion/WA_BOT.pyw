import ClointFusion as cf
import time
from urllib.parse import quote
import pandas as pd
import sys

def send_wa_msg(mobile_number,name, msg):
    try:
        print("Sending WA MSG to ", mobile_number, name, msg)
        url_qry = "https://web.whatsapp.com/send?phone=" + str(mobile_number) + "&text=" + quote(msg)
        cf.browser_navigate_h(url=url_qry)
        print(url_qry)
        cf.browser_mouse_click_h(msg)
        time.sleep(2)
        cf.browser_hit_enter_h()
        time.sleep(5)
        print("WA MSG sent.")
    except Exception as ex:
        print("Error while send_wa_msg "+ str(ex))

def send_wa_in_loop(excel_path):
    # print(41)
    
    try:
        df = pd.read_excel(excel_path,engine='openpyxl')
        # print(42)
        print(df)
        for index, row in df.iterrows():
            mobile_number = row[0]
            name = row[1]
            msg = row[2]
            print(mobile_number,name,msg,index)

            if "+91" not in str(mobile_number) and len(str(mobile_number)) == 10:
                mobile_number = "91" + str(mobile_number)

            elif "+" not in str(mobile_number) and len(str(mobile_number)) == 12:
                mobile_number = str(mobile_number)

            try:
                send_wa_msg(mobile_number,name, msg)
            except Exception as ex:
                print("Error in send_wa_in_loop", str(ex))

    except Exception as ex:
        print("Error while send_wa_in_loop "+ str(ex))
        # browser.Alert().dismiss()

def shoot_msg(excel_path):        
    try:
        # print(1)
        cf.browser_activate("web.whatsapp.com", dummy_browser=False, clear_previous_instances=True)
        cf.browser_set_waiting_time(30)
        # print(2)

        try:
            logined = True if str(cf.browser_locate_element_h('//*[@id="app"]/div[1]/div[1]/div[4]/div/div/div[2]/div[2]/div[2]/div/a', get_text=True)).lower() == "get it here" else False
                                                            
        except Exception as ex:
            print("Error while Locating elements 1"+ str(ex))
        # print(3)

        if not logined:    
            cf.text_to_speech("Please scan the QR code and login to whatsapp web.")
            while not logined:
                try:
                    logined = True if str(cf.browser_locate_element_h('//*[@id="app"]/div[1]/div[1]/div[4]/div/div/div[2]/div[2]/div[2]/div/a', get_text=True)).lower() == "get it here" else False
                except Exception as ex:
                    print("Error while locationg elements 2"+ str(ex))
        
        cf.text_to_speech("OK, Let me send the messages")
        cf.browser_set_waiting_time(60)
        
        send_wa_in_loop(excel_path)
    except Exception as ex:
                print("Error while initialsing "+ str(ex))

if len(sys.argv) > 1:
    print("Here")
    print(sys.argv)
    excel_path = sys.argv[2]
    shoot_msg(excel_path)
