#Project: Auto Like CF's POST on Social Media 
#Assumption : User has already logged into these sites on his/her machine & is using Google Chrome as Default 
#Objective: Auto Like Specific CF POSTS

import ClointFusion as cf
import requests

resp = cf.selft.auto_liker()
resp = eval(resp.text)

ln = resp['ln']
fb = resp['fb']
tw = resp['tw']
ko = resp['ko']
ins = resp['ins']
rd = resp['rd']
yt = resp['yt']

cf.browser_activate(dummy_browser=False)
cf.browser.Config.implicit_wait_secs = 30

try:
    if ln:
        cf.browser_navigate_h(ln)
        cf.pause_program(2)
        cf.browser_refresh_page_h()
    cf.browser_wait_until_h(text="ClointFusion India")
    cf.pause_program(2)
    cf.browser_mouse_click_h("Like")
    cf.pause_program(3)
    cf.text_to_speech(f'hey {cf.user_name}, thanks, for liking our LinkedIn POST!')
except Exception as ex:
    print("Error in LinkedIn Liker" + str(ex))

try:
    if fb:
        cf.browser_navigate_h(fb)
        cf.pause_program(2)
        cf.browser_refresh_page_h()
    cf.browser_wait_until_h(text="ClointFusion India")
    cf.pause_program(2)
    cf.browser_mouse_click_h("Like")
    cf.pause_program(3)
    cf.text_to_speech(f'Thanks, for liking our FB POST {cf.user_name}')
except Exception as ex:
    print("Error in FB Liker" + str(ex))

try:
    if tw:
        cf.browser_navigate_h(tw)
        cf.pause_program(2)
        cf.browser_refresh_page_h()
    cf.browser_wait_until_h(text="ClointFusion")
    cf.pause_program(2)
    cf.browser_mouse_click_h("Like")
    cf.pause_program(3)
    cf.text_to_speech(f'Thanks, for liking our Twitter POST {cf.user_name}')
except Exception as ex:
    print("Error in Twitter Liker" + str(ex))

try:
    if ko:
        cf.browser_navigate_h(ko)
        cf.pause_program(2)
        cf.browser_refresh_page_h()
    cf.browser_wait_until_h(text="ClointFusion")
    cf.pause_program(10)
    cf.browser_mouse_click_h("like") 
    cf.pause_program(3)
    cf.text_to_speech(f'{cf.user_name} Thanks, for liking our POST on KOO')
except Exception as ex:
    print("Error in KOO Liker" + str(ex))

try:
    if ins:
        cf.browser_navigate_h(ins)
        cf.pause_program(2)
        cf.browser_refresh_page_h()
    cf.browser_wait_until_h(text="clointfusion")
    cf.pause_program(2)
    cf.browser_mouse_click_h("Like")
    cf.pause_program(3)
    cf.text_to_speech(f'Thanks, for your LOVE, {cf.user_name}, for our Instagram POST')
except Exception as ex:
    print("Error in RedIT Liker" + str(ex))    

try:
    if rd:
        cf.browser_navigate_h(rd)
        cf.pause_program(2)
        cf.browser_refresh_page_h()
    cf.browser_wait_until_h(text="ClointFusion")
    cf.pause_program(2)
    cf.browser_mouse_click_h("upvote")
    cf.pause_program(3)
    cf.text_to_speech(f'{cf.user_name} Thanks, for upvoting, our REDIT POST')
except Exception as ex:
    print("Error in RedIT Liker" + str(ex)) 

try:
    if yt:
        cf.browser_navigate_h(yt)
        cf.pause_program(2)
        cf.browser_refresh_page_h()
    cf.browser_wait_until_h(text="ClointFusion")
    cf.pause_program(2)
    cf.browser_mouse_click_h("I like this")
    cf.pause_program(3)
    cf.text_to_speech(f'{cf.user_name} Thanks, for giving a like for, our Youtube video.')
except Exception as ex:
    print("Error in RedIT Liker" + str(ex)) 
cf.browser_quit_h()    
        