import os
import random
import requests
import datetime
import sys
import time
import wikipedia
import pyttsx3
import speech_recognition as sr
import pywhatkit as kit
import ClointFusion as cf
from pathlib import Path

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
engine.setProperty('voice', voices[0].id)
r = sr.Recognizer()
energy_threshold = [3000]

if cf.os_name == "windows":
    clointfusion_directory = r"C:\Users\{}\ClointFusion".format(str(os.getlogin()))
elif cf.os_name == "linux":
    clointfusion_directory = r"/home/{}/ClointFusion".format(str(os.getlogin()))
elif cf.os_name == "darwin":
    clointfusion_directory = r"/Users/{}/ClointFusion".format(str(os.getlogin()))

def speak(audio):
    engine.say(audio)   
    engine.runAndWait()

def command():
    print("Listening...")
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
    speak("OK...")
    speak("Which video ?")
    video_name = command().lower() ## takes user command 
    speak("Opening YouTube now, please wait a moment...")
    kit.playonyt(video_name)

def send_WA_MSG():
    speak("OK...")
    speak("Whats the message ?")
    msg = command().lower() ## takes user command 
    if msg not in ["exit", "cancel", "stop"]:
        speak("Got it, whom to send, please say mobile number without country code")
    else:
        speak("Sending message is cancelled...")
        return
    num = command().lower() ## takes user command
    if num not in ["exit", "cancel", "stop"]:
        speak("Sending message now, please wait a moment")
        
        kit.sendwhatmsg_instantly(phone_no=f"+91{num}",message=str(msg),wait_time=25, tab_close=True, close_time=5)
    else:
        speak("Sending message is cancelled...")
        return

def google_search():
    speak("OK...")
    speak("What to search ?")
    msg = command().lower() ## takes user command 
    speak("Searching in Gooogle now, please wait a moment...")

    kit.search(msg)

def greet_user():
    hour = datetime.datetime.now().hour
    greeting = "Good Morning...." if 5<=hour<12 else "Good Afternoon....." if hour<18 else "Good Evening...."
    choices = ["Hey...", "Hi...", "Hello...", "Dear..."]
    greeting = random.choice(choices) + str(cf.user_name) + "..." + greeting + "..."
    speak(greeting + "How can i assist you ?...")
    queries = ["my name..","current time..","global news..","send whatsapp to someone","Send gmail..", "play youtube video...","search in google..."]
    speak("You can ask..")
    choices=random.sample(queries,len(queries))
    speak(choices)
    speak('To quit, you can say...bye...quit...exit...')

def options():
    queries = ["current time..","global news...","send whatsapp to someone...","Send gmail..", "play youtube video...","search in google..."]
    speak("Try saying...")
    choices=random.sample(queries,len(queries))
    speak(choices)
    speak('To quit, you can say...bye...exit...quit...')

def trndnews(): 
    url = "http://newsapi.org/v2/top-headlines?country=in&apiKey=59ff055b7c754a10a1f8afb4583ef1ab"
    page = requests.get(url).json()
    article = page["articles"]
    results = [ar["title"] for ar in article]
    for i in range(len(results)): 
        print(i + 1, results[i])
    speak("Here are the top trending news....!!")
    speak("Do yo want me to read!!!")
    reply = command().lower()
    reply = str(reply)
    if reply == "yes":
        speak(results)
    else:
        speak('ok!!!!')

def bol_main():
    qurey_no = 5

    while True:
        query = command().lower() ## takes user command 
        print("Listening...")
        
        if 'name' in query:
            speak("I am ClointFusion's BOL....")

        ### time
        elif 'time' in query:
            time = datetime.datetime.now().strftime('%I:%M %p')
            speak("It's " + str(time))
        
        ### celebrity
        elif 'who is' in query:
            try:
                query = query.replace('who is',"")
                speak(wikipedia.summary(query,2))
            except:
                speak("Please use a complete word...")

        #Send WA MSG
        elif any(x in query for x in ["send whatsapp","whatsapp","whatsapp message"]): 
            try:
                send_WA_MSG()
            except:
                speak("Sorry, i am experiencing some issues, please try later...")

        #Play YouTube Video
        elif any(x in query for x in ["youtube","play video","video song","youtube video"]): 
            try:
                play_on_youtube()
            except:
                speak("Sorry, i am experiencing some issues, please try later...")

        #Search in Google
        elif any(x in query for x in ["google search","search in google"]): 
            try:
                google_search()
            except:
                speak("Sorry, i am experiencing some issues, please try later...")

        ### news
        elif 'news' in query:
                trndnews() 

        #Clap
        elif any(x in query for x in ["clap","applause","shout","whistle"]):
            _play_sound((str(Path(os.path.join(clointfusion_directory,"Logo_Icons","Applause.wav")))))

        elif any(x in query for x in ["bye","quit","stop","exit"]):
            speak("Have a nice day ! ")
            break

        elif "dost" in query:
            try:
                cf.browser_activate('http://dost.clointfusion.com')
            except:
                pass

        elif "close google chrome" in query:
            try:
                cf.browser_quit_h()
            except:
                pass

        else:
            qurey_no += 1
            
            if qurey_no % 6 == 1:
                options()

greet_user()
bol_main()