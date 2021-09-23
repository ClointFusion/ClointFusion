import os
import random
import webbrowser
import requests
import datetime
import sys
import time
import wikipedia
import pyttsx3
import speech_recognition as sr
import pywhatkit as kit
from . import ClointFusion as cf
from pathlib import Path
import subprocess
from PIL import Image
from rich.text import Text
from rich import print
from rich.console import Console

#Install pyaudio
try:
    import pyaudio
except:
    sys_version = str(sys.version[0:6]).strip()
    
    if "3.7" in sys_version :
        cmd = "https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/Wheels/PyAudio-0.2.11-cp37-cp37m-win_amd64.whl?raw=true"
    elif "3.8" in sys_version :
        cmd = "https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/Wheels/PyAudio-0.2.11-cp38-cp38-win_amd64.whl?raw=true"
    elif "3.9" in sys_version :
        cmd = "https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/Wheels/PyAudio-0.2.11-cp39-cp39-win_amd64.whl?raw=true"

    time.sleep(5)

    try:
        os.system("pip install " + cmd)
    except:
        print("Please install appropriate driver from : https://github.com/ClointFusion/Image_ICONS_GIFs/tree/main/Wheels")

    import pyaudio

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
voice_male_female = random.randint(0,1) # Randomly decide male/female voice
engine.setProperty('voice', voices[voice_male_female].id)
r = sr.Recognizer()
energy_threshold = [3000]
console = Console()

queries = ["current time,","global news,","send whatsapp,","open , minimize , close any application,","Open Gmail,", "play youtube video,","search in google,",'launch zoom meeting,','switch window,','locate on screen,','take selfie,','OCR now,']
latest_queries = ['launch zoom meeting,','switch window,','locate on screen,','take selfie,','OCR now,']

if cf.os_name == "windows":
    clointfusion_directory = r"C:\Users\{}\ClointFusion".format(str(os.getlogin()))
elif cf.os_name == "linux":
    clointfusion_directory = r"/home/{}/ClointFusion".format(str(os.getlogin()))
elif cf.os_name == "darwin":
    clointfusion_directory = r"/Users/{}/ClointFusion".format(str(os.getlogin()))

def text_to_speech(audio):
    if type(audio) is list:
        print(' '.join(audio))
    else:
        print(str(audio))

    engine.say(audio)   
    engine.runAndWait()

def error_try_later():
    text_to_speech("Sorry, i am experiencing some issues, please try later...")

def shuffle_return_one_option(lst_options=[]):
    random.shuffle(lst_options)
    return str(lst_options[0])

def speech_to_text():
    # cf.message_pop_up("listening",1)
    while True:
        with sr.Microphone() as source:
            r.dynamic_energy_threshold = True
            if r.energy_threshold in energy_threshold or r.energy_threshold <= sorted(energy_threshold)[-1]:
                r.energy_threshold = sorted(energy_threshold)[-1]
            else:
                energy_threshold.append(r.energy_threshold)

            r.pause_threshold = 0.6

            r.adjust_for_ambient_noise(source)

            audio=r.listen(source)
            try:
                query = r.recognize_google(audio)
                print(f"You Said : {query}")
                
                cf.clear_screen()
                print("ClointFusion Bol is here to help you")
                return query
                break
            except sr.UnknownValueError:
                pass
            except sr.RequestError as e:
                print("Try Again")

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
        print("Unable to play sound" + str(ex))

def play_on_youtube():
    text_to_speech("OK...")
    text_to_speech("Which video ?")
    video_name = speech_to_text().lower() ## takes user speech_to_text 
    text_to_speech("Opening YouTube now, please wait a moment...")
    kit.playonyt(video_name)

def send_WA_MSG():
    text_to_speech("OK...")
    text_to_speech("Whats the message ?")
    msg = speech_to_text().lower() ## takes user speech_to_text 
    if msg not in ["exit", "cancel", "stop"]:
        text_to_speech("Got it, whom to send, please say mobile number without country code")
    else:
        text_to_speech("Sending message is cancelled...")
        return
    num = speech_to_text().lower() ## takes user speech_to_text
    if num not in ["exit", "cancel", "stop"]:
        text_to_speech("Sending message now, please wait a moment")
        
        kit.sendwhatmsg_instantly(phone_no=f"+91{num}",message=str(msg),wait_time=25, tab_close=True, close_time=5)
    else:
        text_to_speech("Sending message is cancelled...")
        return

def google_search():
    text_to_speech("OK...")
    text_to_speech("What to search ?")
    msg = speech_to_text().lower() ## takes user speech_to_text 
    text_to_speech("Searching in Gooogle now, please wait a moment...")

    kit.search(msg)

def welcome(nth):
    hour = datetime.datetime.now().hour
    greeting = " Good Morning ! " if 5<=hour<12 else " Good Afternoon !! " if hour<18 else " Good Evening !"
    choices = ["Hey ! ", "Hi ! ", "Hello ! ", "Dear ! "]
    greeting = random.choice(choices) + str(cf.user_name) + ' !!' + greeting
    text_to_speech(greeting + " How can i assist you ?!")
    
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
    text_to_speech("Try saying...")
    random.shuffle(queries)
    
    text_to_speech(queries)
    quit_options=['bye','quit','exit']
    random.shuffle(quit_options)
    text_to_speech('To quit, just say {}'.format(quit_options[0]))

def help():
    text_to_speech("Here are all the commands i support currently.")
    print("All commands")
    print(queries)
    text_to_speech("Try some latest commands.")
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
    text_to_speech("Here are the top trending news....!!")
    text_to_speech("Do yo want me to read!!!")
    reply = speech_to_text().lower()
    reply = str(reply)
    if reply == "yes":
        text_to_speech(results)
    else:
        text_to_speech('ok!!!!')

def capture_photo(ocr=False):
    try:
        subprocess.run('start microsoft.windows.camera:', shell=True)

        if ocr:
            time.sleep(4)
        else:
            time.sleep(1)

        img=cf.pg.screenshot()
        time.sleep(1)
        
        img.save(Path(os.path.join(clointfusion_directory, "Images","Selfie.PNG")))                    
        subprocess.run('Taskkill /IM WindowsCamera.exe /F', shell=True)
    except Exception as ex:
        print("Error in capture_photo " + str(ex))

def bol_main():
    query_num = 5
    with console.status("Listening...\n") as status:
        while True:
            query = speech_to_text().lower() ## takes user speech_to_text
            
            if 'name' in query:
                text_to_speech("I am ClointFusion's BOL....")

            ### time
            elif 'time' in query:
                time = datetime.datetime.now().strftime('%I:%M %p')
                text_to_speech("It's " + str(time))
            
            elif any(x in query for x in ["help","commands", "list of commands", "what can you do",]):
                help()
            
            ### celebrity
            elif 'who is' in query:
                try:
                    query = query.replace('who is',"")
                    text_to_speech(wikipedia.summary(query,2))
                except:
                    text_to_speech("Please use a complete word...")

            #Send WA MSG
            elif any(x in query for x in ["send whatsapp","whatsapp","whatsapp message"]): 
                try:
                    send_WA_MSG()
                except:
                    error_try_later()

            #Play YouTube Video
            elif any(x in query for x in ["youtube","play video","video song","youtube video"]): 
                try:
                    play_on_youtube()
                except:
                    error_try_later()

            #Search in Google
            elif any(x in query for x in ["google search","search in google"]): 
                try:
                    google_search()
                except:
                    error_try_later()

            #Open gmail
            elif any(x in query for x in ["gmail","email"]): 
                try:
                    webbrowser.open_new_tab("http://mail.google.com")
                except:
                    error_try_later()

            #open camera
            elif any(x in query for x in ["launch camera","open camera"]): 
                try:
                    subprocess.run('start microsoft.windows.camera:', shell=True)
                except:
                    os.startfile('microsoft.windows.camera:')

            ### close camera
            elif any(x in query for x in ["close camera"]): 
                subprocess.run('Taskkill /IM WindowsCamera.exe /F', shell=True)

            ### news
            elif 'news' in query:
                trndnews() 

            #Clap
            elif any(x in query for x in ["clap","applause","shout","whistle"]):
                _play_sound((str(Path(os.path.join(clointfusion_directory,"Logo_Icons","Applause.wav")))))

            elif any(x in query for x in ["bye","quit","stop","exit"]):
                exit_say_choices=["Have a good day! ","Have an awesome day!","I hope your day is great!","Today will be the best!","Have a splendid day!","Have a nice day!","Have a pleasant day!"]
                text_to_speech(shuffle_return_one_option(exit_say_choices))
                break

            elif "dost" in query:
                try:
                    cf.browser_activate('http://dost.clointfusion.com')
                except:
                    pass

            elif any(x in query for x in ["open notepad","launch notepad"]):
                try:
                    cf.launch_any_exe_bat_application("notepad")
                except:
                    pass

            elif any(x in query for x in ["open application","launch application","launch app","open app"]):
                try:
                    text_to_speech('OK, which application to open?')
                    app_name = speech_to_text().lower() ## takes user speech_to_text 
                    cf.launch_any_exe_bat_application(app_name)
                except:
                    pass

            #Switch to window
            elif any(x in query for x in ["switch window","toggle window","activate window","maximize window"]): 
                try:
                    text_to_speech('OK, whats the window name?')
                    windw_name = speech_to_text().lower() ## takes user speech_to_text 
                    cf.window_activate_and_maximize_windows(windw_name)
                except:
                    error_try_later()

            #Search in window / browser
            elif any(x in query for x in ["find on screen","search on screen", "locate on screen"]): 
                try:
                    text_to_speech('OK, what to find ?')
                    query = speech_to_text().lower() ## takes user speech_to_text 
                    cf.find_text_on_screen(searchText=query,delay=0.1, occurance=1,isSearchToBeCleared=False)
                except:
                    error_try_later()

            elif any(x in query for x in ["minimize all","minimize window","show desktop"]):
                try:
                    cf.window_show_desktop()
                except:
                    pass              

            elif any(x in query for x in ["minimize window","minimize application"]):
                try:
                    text_to_speech('OK, which window to minimize?')
                    windw_name = speech_to_text().lower() ## takes user speech_to_text 
                    cf.window_minimize_windows(windw_name)
                except:
                    pass 

            elif any(x in query for x in ["close application","close window"]):
                try:
                    text_to_speech('OK, which application to close?')
                    app_name = speech_to_text().lower() ## takes user speech_to_text 
                    cf.window_close_windows(app_name)
                except:
                    pass

            elif any(x in query for x in ["launch meeting","zoom meeting"]):
                try:
                    webbrowser.open_new_tab("https://us02web.zoom.us/j/85905538540?pwd=b0ZaV3c2bC9zK3I1QXNjYjJ3Q0tGdz09")
                except:
                    pass

            elif "close google chrome" in query:
                try:
                    cf.browser_quit_h()
                except:
                    pass

            elif any(x in query for x in ["take pic","take selfie","take a pic","take a selfie"]):
                try:
                    smile_say=["OK, Smile Please !","OK, Please look at the Camera !","OK, Say Cheese !","OK, Sit up straight !"]

                    text_to_speech(shuffle_return_one_option(smile_say))

                    capture_photo()
                    text_to_speech("Thanks, I saved your photo. Do you want me to open ?")
                    yes_no = speech_to_text().lower() ## takes user speech_to_text 
                    if yes_no in ["yes", "yah", "ok"]:
                        cf.launch_any_exe_bat_application(Path(os.path.join(clointfusion_directory, "Images","Selfie.PNG")))
                    else:
                        pass
                except:
                    error_try_later()

            elif any(x in query for x in ["clear screen","clear","clear terminal","clean", "clean terminal","clean screen",]):
                try:
                    cf.clear_screen()
                    print("ClointFusion Bol is here to help.")
                except Exception as ex:
                    print("Error in clearing " + str(ex))
                    error_try_later()
            
            elif 'ocr' in query:
                try:
                    ocr_say=["OK, Let me scan !","OK, Going to scan now","Please show me the image"]

                    text_to_speech(shuffle_return_one_option(ocr_say))

                    capture_photo(ocr=True)

                    ocr_img_path = Path(os.path.join(clointfusion_directory, "Images","Selfie.PNG"))

                    imageObject = Image.open(ocr_img_path)

                    corrected_image = imageObject.transpose(Image.FLIP_LEFT_RIGHT)

                    corrected_image.save(ocr_img_path)

                    text_to_speech(shuffle_return_one_option(["OK, performing OCR now","Give me a moment","abracadabra","Hang on"]))

                    ocr_result = cf.ocr_now(ocr_img_path)
                    print(ocr_result)
                    text_to_speech(ocr_result)

                except Exception as ex:
                    print("Error in OCR " + str(ex))
                    error_try_later()

            elif any(x in query for x in ["thanks","thank you"]):
                choices = ["You're welcome","You're very welcome.","That's all right.","No problem.","No worries.","Don't mention it.","It's my pleasure.","My pleasure.","Glad to help.","Sure!",""]
                
                text_to_speech(shuffle_return_one_option(choices))            

            elif any(x in query for x in ["shutdown my","turn off","switch off"]):
                try:
                    text_to_speech('Do you want to Shutdown ? Are you sure ?')
                    yes_no = speech_to_text().lower() ## takes user speech_to_text 
                    if yes_no in ["yes", "yah", "ok"]:
                        text_to_speech("OK, Shutting down your machine")

                        os.system('shutdown -s')
                    
                except:
                    error_try_later()

            else:
                query_num += 1
            
                if query_num % 6 == 1:
                    options()

bol_config = clointfusion_directory + "\Config_Files\_bol.txt"

try:
    if os.path.exists(bol_config):
        with open(bol_config, 'r') as fp:
            nth = fp.readline(1)
        with open(bol_config, 'w') as fp:
            fp.write(str(int(nth) + 1))
        welcome(int(nth))
    else:
        with open(bol_config, 'w') as fp:
            fp.write("1")
            welcome(1)
except:
    pass


bol_main()
