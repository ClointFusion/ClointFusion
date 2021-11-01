import os, webbrowser, requests, time, sys, sqlite3, random, traceback
import pyinspect as pi
import datetime
import wikipedia
import pywhatkit as kit
import ClointFusion as cf
from pathlib import Path
import subprocess
from PIL import Image
from rich.text import Text
from rich import print
from rich.console import Console
import pyaudio
from rich import pretty
import platform

windows_os = "windows"
linux_os = "linux"
mac_os = "darwin"
os_name = str(platform.system()).lower()

pi.install_traceback(hide_locals=True,relevant_only=True,enable_prompt=True)
pretty.install()

console = Console()

queries = ["current time,","global news,","send whatsapp,","open , minimize , close any application,","Open Gmail,", "play youtube video,","search in google,",'launch zoom meeting,','switch window,','locate on screen,','take selfie,','OCR now,', 'commands,', 'read screen,','help,',]
latest_queries = ['launch zoom meeting,','switch window,','locate on screen,','take selfie,','read screen,',]

def error_try_later():
    error_choices=["Whoops, please try again","Mea Culpa, please try again","Sorry, i am experiencing some issues,  please try again","Apologies, please try again"]
    cf.text_to_speech(shuffle_return_one_option(error_choices))

def shuffle_return_one_option(lst_options=[]):
    random.shuffle(lst_options)
    return str(lst_options[0])

def _play_sound(music_file_path=""):
    try:
        import wave  
        #define stream chunk   
        chunk = 1024  

        #open a wav format music  
        f = wave.open(music_file_path,"rb")  
        #instantiate PyAudio  
        p = pyaudio.PyAudio()  
        #open stream  
        stream = p.open(format = p.get_format_from_width(f.getsampwidth()),  
                        channels = f.getnchannels(),  
                        rate = f.getframerate(),  
                        output = True)  
        
        data = f.readframes(chunk)  

        #play stream  
        while data:  
            stream.write(data)  
            data = f.readframes(chunk)  

        #stop stream  
        stream.stop_stream()  
        stream.close()  

        #close PyAudio  
        p.terminate()
    except Exception as ex:
        cf.selft.crash_report(traceback.format_exception(*sys.exc_info(),limit=None, chain=True))
        print("Unable to play sound" + str(ex))

def play_on_youtube():
    cf.text_to_speech("OK...")
    cf.text_to_speech("Which video ?")
    video_name = cf.speech_to_text().lower() ## takes user cf.speech_to_text 
    cf.text_to_speech("Opening YouTube now, please wait a moment...")
    kit.playonyt(video_name)

def call_Send_WA_MSG():
    cf.text_to_speech("OK...")
    cf.text_to_speech("Whats the message ?")
    msg = cf.speech_to_text().lower() ## takes user cf.speech_to_text 
    if msg not in ["exit", "cancel", "stop"]:
        cf.text_to_speech("Got it, whom to send, please say mobile number without country code")
    else:
        cf.text_to_speech("Sending message is cancelled...")
        return
    num = cf.speech_to_text().lower() ## takes user cf.speech_to_text
    if num not in ["exit", "cancel", "stop"]:
        cf.text_to_speech("Sending message now, please wait a moment")
        
        kit.sendwhatmsg_instantly(phone_no=f"+91{num}",message=str(msg),wait_time=25, tab_close=True, close_time=5)
    else:
        cf.text_to_speech("Sending message is cancelled...")
        return

def google_search():
    cf.text_to_speech("OK...")
    cf.text_to_speech("What to search ?")
    msg = cf.speech_to_text().lower() ## takes user cf.speech_to_text 
    cf.text_to_speech("Searching in Gooogle now, please wait a moment...")

    kit.search(msg)

def welcome(nth):
    hour = datetime.datetime.now().hour
    greeting = " Good Morning ! " if 5<=hour<12 else " Good Afternoon !! " if hour<18 else " Good Evening ! "
    choices = ["Hey ! ", "Hi ! ", "Hello ! ", "Dear ! "]
    greeting = random.choice(choices) + str(cf.user_name) + ' !!' + greeting
    assist_choices = ["How can i assist you ?!","What can I do for you?","Feel free to call me for any help!","What something can I do for you?"]
    cf.text_to_speech(greeting + shuffle_return_one_option(assist_choices))
    
    if nth == 1:
        suggest(options(3, 2))
    elif nth == 2:
        suggest(options(5, 5))
    elif nth == 3:
        suggest(options(5, 1))
    else:
        pass
    
def suggest(suggestions):
    queries = suggestions
    cf.text_to_speech("Try saying...")
    random.shuffle(queries)
    
    cf.text_to_speech(queries)
    quit_options=['bye','quit','exit']
    random.shuffle(quit_options)
    cf.text_to_speech('To quit, just say {}'.format(quit_options[0]))

def call_help():
    cf.text_to_speech(shuffle_return_one_option(["I support these commands","Here are the commands i support currently."]))
    print("All commands:")
    print(queries)
    cf.text_to_speech(shuffle_return_one_option(["Try some latest commands:","Try something new:"]))
    print("Latest commands")
    print(latest_queries)
    print("\n")

def options(total=5, latest=3):
    remaining = [q for q in queries if q not in latest_queries]
    custom_list = []
    latest_done = False
    try :
      for i in range(latest+1):
        custom_list.append(latest_queries[i])
      latest_done = True
    except IndexError:
      latest_done = True
    if latest_done:
      for i in range(total - len(custom_list)):
        custom_list.append(remaining[i])
    random.shuffle(custom_list)
    return custom_list

def trndnews(): 
    url = "http://newsapi.org/v2/top-headlines?country=in&apiKey=59ff055b7c754a10a1f8afb4583ef1ab"
    page = requests.get(url).json()
    article = page["articles"]
    results = [ar["title"] for ar in article]
    for i in range(len(results)): 
        print(i + 1, results[i])
    cf.text_to_speech("Here are the top trending news....!!")
    cf.text_to_speech("Do yo want me to read!!!")
    reply = cf.speech_to_text().lower()
    reply = str(reply)
    if reply == "yes":
        cf.text_to_speech(results)
    else:
        cf.text_to_speech('ok!!!!')

def capture_photo(ocr=False):
    try:
        subprocess.run('start microsoft.windows.camera:', shell=True)

        if ocr:
            time.sleep(4)
        else:
            time.sleep(1)

        img=cf.pg.screenshot()
        time.sleep(1)
        
        img.save(Path(os.path.join(cf.clointfusion_directory, "Images","Selfie.PNG")))                    
        subprocess.run('Taskkill /IM WindowsCamera.exe /F', shell=True)
    except Exception as ex:
        cf.selft.crash_report(traceback.format_exception(*sys.exc_info(),limit=None, chain=True))
        print("Error in capture_photo " + str(ex))

def call_read_screen():
    try:
        cf.text_to_speech('Window Name to read?')
        windw_name = cf.speech_to_text().lower() ## takes user cf.speech_to_text 
        cf.window_show_desktop()
        cf.window_activate_and_maximize_windows(windw_name)
        time.sleep(2)
        img=cf.pg.screenshot()
        img.save(Path(os.path.join(cf.clointfusion_directory, "Images","Selfie.PNG")))
        # OCR process
        ocr_img_path = Path(os.path.join(cf.clointfusion_directory, "Images","Selfie.PNG"))
        cf.text_to_speech(shuffle_return_one_option(["OK, performing OCR now","Give me a moment","abracadabra","Hang on"]))
        ocr_result = cf.ocr_now(ocr_img_path)
        print(ocr_result)

        cf.text_to_speech("Do you want me to read?")
        yes_no = cf.speech_to_text().lower() ## takes user cf.speech_to_text 
        if yes_no in ["yes", "yah", "ok"]:
            cf.text_to_speech(ocr_result)
    except Exception as ex:
        cf.selft.crash_report(traceback.format_exception(*sys.exc_info(),limit=None, chain=True))
        print("Error in capture_photo " + str(ex))

def call_name():
    name_choices = ["I am ClointFusion's BOL!", "This is Bol!", "Hi, I am ClointFusion's Bol!","Hey, this is Bol!"]
    cf.text_to_speech(shuffle_return_one_option(name_choices))

def call_time():
    time = datetime.datetime.now().strftime('%I:%M %p')
    cf.text_to_speech("It's " + str(time))

def call_wiki(query):
    try:
        cf.text_to_speech(wikipedia.summary(query,2))
    except:
        cf.text_to_speech("Please use a complete word...")

def call_ocr():
    try:
        ocr_say=["OK, Let me scan !","OK, Going to scan now","Please show me the image"]

        cf.text_to_speech(shuffle_return_one_option(ocr_say))

        capture_photo(ocr=True)

        ocr_img_path = Path(os.path.join(cf.clointfusion_directory, "Images","Selfie.PNG"))

        imageObject = Image.open(ocr_img_path)

        corrected_image = imageObject.transpose(Image.FLIP_LEFT_RIGHT)

        corrected_image.save(ocr_img_path)

        cf.text_to_speech(shuffle_return_one_option(["OK, performing OCR now","Give me a moment","abracadabra","Hang on"]))

        ocr_result = cf.ocr_now(ocr_img_path)
        print(ocr_result)
        cf.text_to_speech(ocr_result)

    except Exception as ex:
        cf.selft.crash_report(traceback.format_exception(*sys.exc_info(),limit=None, chain=True))
        print("Error in OCR " + str(ex))
        error_try_later()

def call_camera():
    try:
        subprocess.run('start microsoft.windows.camera:', shell=True)
    except:
        os.startfile('microsoft.windows.camera:')

def call_any_app():
    cf.text_to_speech('OK, which application to open?')
    app_name = cf.speech_to_text().lower() ## takes user cf.speech_to_text 
    cf.launch_any_exe_bat_application(app_name)

def call_switch_wndw():
    cf.text_to_speech('OK, whats the window name?')
    windw_name = cf.speech_to_text().lower() ## takes user cf.speech_to_text 
    cf.window_activate_and_maximize_windows(windw_name)

def call_find_on_screen():
    cf.text_to_speech('OK, what to find ?')
    query = cf.speech_to_text().lower() ## takes user cf.speech_to_text 
    cf.find_text_on_screen(searchText=query,delay=0.1, occurance=1,isSearchToBeCleared=False)

def call_minimize_wndw():
    cf.text_to_speech('OK, which window to minimize?')
    windw_name = cf.speech_to_text().lower() ## takes user cf.speech_to_text 
    cf.window_minimize_windows(windw_name)

def call_close_app():
    cf.text_to_speech('OK, which application to close?')
    app_name = cf.speech_to_text().lower() ## takes user cf.speech_to_text 
    cf.window_close_windows(app_name)

def call_take_selfie():
    smile_say=["OK, Smile Please !","OK, Please look at the Camera !","OK, Say Cheese !","OK, Sit up straight !"]

    cf.text_to_speech(shuffle_return_one_option(smile_say))

    capture_photo()
    cf.text_to_speech("Thanks, I saved your photo. Do you want me to open ?")
    yes_no = cf.speech_to_text().lower() ## takes user cf.speech_to_text 
    if yes_no in ["yes", "yah", "ok"]:
        cf.launch_any_exe_bat_application(Path(os.path.join(cf.clointfusion_directory, "Images","Selfie.PNG")))
    else:
        pass

def call_thanks():
    choices = ["You're welcome","You're very welcome.","That's all right.","No problem.","No worries.","Don't mention it.","It's my pleasure.","My pleasure.","Glad to help.","Sure!",""]
    cf.text_to_speech(shuffle_return_one_option(choices))   

def call_shut_pc():
    cf.text_to_speech('Do you want to Shutdown ? Are you sure ?')
    yes_no = cf.speech_to_text().lower() ## takes user cf.speech_to_text 
    if yes_no in ["yes", "yah", "ok"]:
        cf.text_to_speech("OK, Shutting down your machine in a minute")
        os.system('shutdown -s')

def call_social_media():
    #opens all social media links of ClointFusion
    try:
        webbrowser.open_new_tab("https://www.facebook.com/ClointFusion")
    except Exception as ex:
        print("Error in call_social_media = " + str(ex))
 
    try:
        webbrowser.open_new_tab("https://twitter.com/ClointFusion")
    except Exception as ex:
        print("Error in call_social_media = " + str(ex))

    try:
        webbrowser.open_new_tab("https://www.youtube.com/channel/UCIygBtp1y_XEnC71znWEW2w")
    except Exception as ex:
        print("Error in call_social_media = " + str(ex))

    try:
        webbrowser.open_new_tab("https://www.linkedin.com/showcase/clointfusion_official")
    except Exception as ex:
        print("Error in call_social_media = " + str(ex))
    try:
        webbrowser.open_new_tab("https://www.reddit.com/user/Cloint-Fusion")
    except Exception as ex:
        print("Error in call_social_media = " + str(ex))

    try:
        webbrowser.open_new_tab("https://www.instagram.com/clointfusion")
    except Exception as ex:
        print("Error in call_social_media = " + str(ex))

    try:
        webbrowser.open_new_tab("https://www.kooapp.com/profile/ClointFusion")
    except Exception as ex:
        print("Error in call_social_media = " + str(ex))

    try:
        webbrowser.open_new_tab("https://discord.com/invite/tsMBN4PXKH")
    except Exception as ex:
        print("Error in call_social_media = " + str(ex))

    try:
        webbrowser.open_new_tab("https://www.eventbrite.com/e/2-days-event-on-software-bot-rpa-development-with-no-coding-tickets-183070046437")
    except Exception as ex:
        print("Error in call_social_media = " + str(ex))

    try:
        webbrowser.open_new_tab("https://internshala.com/internship/detail/python-rpa-automation-software-bot-development-work-from-home-job-internship-at-clointfusion1631715670")
    except Exception as ex:
        print("Error in call_social_media = " + str(ex))

def bol_main():
    query_num = 5
    with console.status("Listening...\n") as status:
        while True:
            query = cf.speech_to_text().lower() ## takes user cf.speech_to_text
            
            try:
                if any(x in query for x in ["name","bol"]):
                    call_name()

                elif 'time' in query:
                    call_time()
                
                elif any(x in query for x in ["help","commands", "list of commands", "what can you do",]):
                    call_help()
                
                elif 'who is' in query:
                    status.update("Processing...\n")
                    query = query.replace('who is',"")
                    call_wiki(query)
                    status.update("Listening...\n")

                #Send WA MSG
                elif any(x in query for x in ["send whatsapp","whatsapp","whatsapp message"]):
                    status.update("Processing...\n")
                    call_Send_WA_MSG()
                    status.update("Listening...\n")
                    
                #Play YouTube Video
                elif any(x in query for x in ["youtube","play video","video song","youtube video"]):
                    status.update("Processing...\n")
                    play_on_youtube()
                    status.update("Listening...\n")
                
                #Search in Google
                elif any(x in query for x in ["google search","search in google"]): 
                    status.update("Processing...\n")
                    google_search()
                    status.update("Listening...\n")
                
                #Open gmail
                elif any(x in query for x in ["gmail","email"]): 
                    status.update("Processing...\n")
                    webbrowser.open_new_tab("http://mail.google.com")
                    status.update("Listening...\n")
                    
                #open camera
                elif any(x in query for x in ["launch camera","open camera"]): 
                    status.update("Processing...\n")
                    call_camera()
                    status.update("Listening...\n")

                ### close camera
                elif any(x in query for x in ["close camera"]): 
                    status.update("Processing...\n")
                    subprocess.run('Taskkill /IM WindowsCamera.exe /F', shell=True)
                    status.update("Listening...\n")

                ### news
                elif 'news' in query:
                    status.update("Processing...\n")
                    trndnews() 
                    status.update("Listening...\n")

                #Clap
                elif any(x in query for x in ["clap","applause","shout","whistle"]):
                    status.update("Processing...\n")
                    _play_sound((str(Path(os.path.join(cf.clointfusion_directory,"Logo_Icons","Applause.wav")))))
                    status.update("Listening...\n")

                elif any(x in query for x in ["bye","quit","stop","exit"]):
                    exit_say_choices=["Have a good day! ","Have an awesome day!","I hope your day is great!","Today will be the best!","Have a splendid day!","Have a nice day!","Have a pleasant day!"]
                    cf.text_to_speech(shuffle_return_one_option(exit_say_choices))
                    break

                elif "dost" in query:
                    try:
                        import subprocess
                        try:
                            import site  
                            site_packages_path = next(p for p in site.getsitepackages() if 'site-packages' in p)
                        except:
                            site_packages_path = subprocess.run('python -c "import os; print(os.path.join(os.path.dirname(os.__file__), \'site-packages\'))"',capture_output=True, text=True).stdout

                        site_packages_path = str(site_packages_path).strip()  

                        status.update("Processing...\n")
                        status.stop()
                        cmd = f'python "{site_packages_path}\ClointFusion\DOST_CLIENT.pyw"'
                        os.system(cmd)
                        status.start()
                        status.update("Listening...\n")
                    except Exception as ex:
                        cf.selft.crash_report(traceback.format_exception(*sys.exc_info(),limit=None, chain=True))
                        print("Error in calling dost from bol = " + str(ex))

                elif any(x in query for x in ["open notepad","launch notepad"]):
                    status.update("Processing...\n")
                    cf.launch_any_exe_bat_application("notepad")
                    status.update("Listening...\n")
                    
                elif any(x in query for x in ["open application","launch application","launch app","open app"]):
                    status.update("Processing...\n")
                    call_any_app()
                    status.update("Listening...\n")
                  
                #Switch to window
                elif any(x in query for x in ["switch window","toggle window","activate window","maximize window"]): 
                    status.update("Processing...\n")
                    call_switch_wndw()
                    status.update("Listening...\n")

                #Search in window / browser
                elif any(x in query for x in ["find on screen","search on screen", "locate on screen"]): 
                    status.update("Processing...\n")
                    call_find_on_screen()
                    status.update("Listening...\n")
                
                elif any(x in query for x in ["minimize all","minimize window","show desktop"]):
                    status.update("Processing...\n")
                    cf.window_show_desktop()
                    status.update("Listening...\n")
                    
                elif any(x in query for x in ["minimize window","minimize application"]):
                    status.update("Processing...\n")
                    call_minimize_wndw()
                    status.update("Listening...\n")

                elif any(x in query for x in ["close application","close window"]):
                    status.update("Processing...\n")
                    call_close_app()
                    status.update("Listening...\n")

                elif any(x in query for x in ["launch meeting","zoom meeting"]):
                    status.update("Processing...\n")
                    webbrowser.open_new_tab("https://us02web.zoom.us/j/85905538540?pwd=b0ZaV3c2bC9zK3I1QXNjYjJ3Q0tGdz09")
                    status.update("Listening...\n")

                elif "close google chrome" in query:
                    status.update("Processing...\n")
                    cf.browser_quit_h()
                    status.update("Listening...\n")

                elif any(x in query for x in ["take pic","take selfie","take a pic","take a selfie"]):
                    status.update("Processing...\n")
                    call_take_selfie()
                    status.update("Listening...\n")

                elif any(x in query for x in ["clear screen","clear","clear terminal","clean", "clean terminal","clean screen",]):
                    status.update("Processing...\n")
                    cf.clear_screen()
                    print("ClointFusion Bol is here to help.")
                
                elif 'ocr' in query:
                    status.update("Processing...\n")
                    call_ocr()

                elif any(x in query for x in ["social media"]):
                    status.update("Processing...\n")
                    call_social_media()
                    status.update("Listening...\n")

                elif any(x in query for x in ["read the screen","read screen","screen to text"]):
                    status.update("Processing...\n")
                    call_read_screen()
                    status.update("Listening...\n")
                
                elif any(x in query for x in ["thanks","thank you"]):
                    status.update("Processing...\n")
                    call_thanks()
                    status.update("Listening...\n")

                elif any(x in query for x in ["shutdown my","turn off","switch off"]):
                    status.update("Processing...\n")
                    call_shut_pc()
                    status.update("Listening...\n")
                    
                else:
                    query_num += 1
                
                    if query_num % 6 == 1:
                        options(3, 2)

            except:
                error_try_later()


config_folder_path = Path(os.path.join(cf.clointfusion_directory, "Config_Files")) 

if os_name == windows_os:
    db_file_path = r'{}\BRE_WHM.db'.format(str(config_folder_path))
else:
    db_file_path = cf.folder_create_text_file(config_folder_path, 'BRE_WHM.db', custom=True)
        
try:
    connct = sqlite3.connect(db_file_path,check_same_thread=False)
    cursr = connct.cursor()
except Exception as ex:
    cf.selft.crash_report(traceback.format_exception(*cf.sys.exc_info(),limit=None, chain=True))
    print("Error in connecting to DB="+str(ex))        


data = cursr.execute("SELECT bol from CF_VALUES where ID = 1")
for row in data:
   run =  row[0]

welcome(run)
cursr.execute("UPDATE CF_VALUES set bol = bol+1 where ID = 1")
connct.commit()

bol_main()
