#Project: Auto Like CF's POST on Social Media 
#Assumption : User has already logged into these sites on his/her machine & is using Google Chrome as Default 
#Objective: Auto Like Specific CF POSTS

import ClointFusion as cf
import requests

resp = requests.post('https://api.clointfusion.com/auto_liker',data={'uuid': str(cf.selft.get_uuid())},)
resp = eval(resp.text)
print(resp)

cf.browser.Config.implicit_wait_secs = 30

try:
    cf.browser_activate(resp['ln'],dummy_browser=False) 
    cf.browser_wait_until_h(text="ClointFusion India")
    cf.pause_program(2)
    cf.browser_mouse_click_h("Like")
    cf.pause_program(3)
    cf.text_to_speech(f'hey {cf.user_name}, thanks, for liking our LinkedIn POST!')
except Exception as ex:
    print("Error in LinkedIn Liker" + str(ex))

try:
    cf.browser_navigate_h(resp['fb'])
    cf.browser_wait_until_h(text="ClointFusion India")
    cf.pause_program(2)
    cf.browser_mouse_click_h("Like")
    cf.pause_program(3)
    cf.text_to_speech(f'Thanks, for liking our FB POST {cf.user_name}')
except Exception as ex:
    print("Error in FB Liker" + str(ex))

try:
    cf.browser_navigate_h(resp['tw'])
    cf.browser_wait_until_h(text="ClointFusion")
    cf.pause_program(2)
    cf.browser_mouse_click_h("Like")
    cf.pause_program(3)
    cf.text_to_speech(f'Thanks, for liking our Twitter POST {cf.user_name}')
except Exception as ex:
    print("Error in Twitter Liker" + str(ex))

try:
    cf.browser_navigate_h(resp['ko']) 
    cf.browser_wait_until_h(text="ClointFusion")
    cf.pause_program(10)
    cf.browser_mouse_click_h("like") 
    cf.pause_program(3)
    cf.text_to_speech(f'{cf.user_name} Thanks, for liking our POST on KOO')
except Exception as ex:
    print("Error in KOO Liker" + str(ex))

try:
    cf.browser_navigate_h(resp['ins'])
    cf.browser_wait_until_h(text="clointfusion")
    cf.pause_program(2)
    cf.browser_mouse_click_h("Like")
    cf.pause_program(3)
    cf.text_to_speech(f'Thanks, for your LOVE, {cf.user_name}, for our Instagram POST')
except Exception as ex:
    print("Error in RedIT Liker" + str(ex))    

try:
    cf.browser_navigate_h(resp['rd'])
    cf.browser_wait_until_h(text="ClointFusion")
    cf.pause_program(2)
    cf.browser_mouse_click_h("upvote")
    cf.pause_program(3)
    cf.text_to_speech(f'{cf.user_name} Thanks, for upvoting, our REDIT POST')
except Exception as ex:
    print("Error in RedIT Liker" + str(ex)) 

cf.browser_quit_h()    
        