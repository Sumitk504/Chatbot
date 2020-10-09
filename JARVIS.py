import pyttsx3
import datetime
import speech_recognition as sr
import pyautogui
import wikipedia
import psutil
import pyjokes
import calendar
import sounddevice as sd
import soundfile as sf
import imdb
import socket
import os
import turtle
import ctypes
import smtplib
import webbrowser as wb
import win32com.client as winc1
from ecapture import ecapture as ec
from speedtest import Speedtest




engine = pyttsx3.init()

# --------------Maine Speak function-----------------
def speak(audio):
    engine.say(audio)
    engine.runAndWait()

# --------------Time function------------------------
def time():
    speak("The current time is")
    Time = datetime.datetime.now().strftime("%H:%M:%S")
    speak(Time)

# --------------Date function------------------------
def date():
    year = int(datetime.datetime.now().year)
    month = int(datetime.datetime.now().month)
    day = int(datetime.datetime.now().day)
    speak(day)
    speak(month)
    speak(year)
    

# --------------wish function------------------------
def wishme():
    hrs = datetime.datetime.now().hour
    if hrs>=6 and hrs<=12:
        speak("Good Morning sir")
    elif hrs>12 and hrs<=18:
        speak("Good After noon sir")
    elif hrs>18 and hrs<=24:
        speak("Good Evening sir")
    else:
        speak("Good night sir")
    speak("How can i help you")

# --------------Screenshot function-------------------
def screenshot():
    img = pyautogui.screenshot()
    img.save("E:\Projects\Jarvis\img\ss.jpeg")

# --------------CPU and Battery update function--------
def cpu():
    usage = str(psutil.cpu_percent())
    speak("cpu is at" + usage)
    battery = psutil.sensors_battery()
    speak("battery is at")
    speak(battery.percent)

# --------------Sound Recording function-----------------
def voice_rec():
    speak("Okay sir, Speak something")
    fs = 48000
    duration = 7 #seconds
    print("Recording")
    myrecording = sd.rec(int(duration * fs), samplerate=fs, channels=2)
    sd.wait()
    return sf.write('my_Audio_Message_file1.flac', myrecording, fs)
    print("Recorded")




def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(query)

    except Exception as e:
        print(e)
        speak("say theat again please")
        return "None"
    return query

if __name__ == "__main__":
    wishme()
    while True:
        query = takeCommand().lower()
        if 'time' in query:
            time()

        # elif 'date' in query:
            # date()

        elif 'open calender' in query:
            cal = calendar.TextCalendar(calendar.SUNDAY)
            for m in range(1,13):
                print(cal.formatmonth(2020,m))

        # ------------------Wikipedia Search-------------------
        elif 'wikipedia' in query:
            speak("Searching...")
            query = query.replace("wikipedia","")
            result = wikipedia.summary(query, sentences = 2)
            print(result)
            speak(result)
        
        # ------------------IMDB rating------------------------
        elif 'rating movies' in query:
            speak("Okay sir!, Here is the top 5 movies")
            print("Okay sir!, Here is the top 5 movies")
            ia = imdb.IMDb()
            search = ia.get_top250_movies()
            for i in range(5):
                print(search[i])
                speak(search[i])
            
        # ------------------Chrome Search-----------------------
        elif 'open chrome' in query:
            speak("What should I search?")
            chromepath = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'
            search = takeCommand().lower()
            wb.get(chromepath).open_new_tab(search+'.com')
        
        # ------------------Music Player-----------------------
        elif 'play some music' in query:
            speak("okay sir")
            songs_dir = 'D:\\Download'
            songs = os.listdir(songs_dir)
            os.startfile(os.path.join(songs_dir, songs[5]))
        
        # ---------------Notes Saver and NOtes Teller------------
        elif 'write a note' in query:
            speak("What should i write")
            data = takeCommand()
            speak("You said me to remember that" +data)
            remember = open('data.txt', 'w')
            remember.write(data)
            remember.close()

        elif 'do you know anything' in query:
            remember = open('data.txt', 'r')
            speak("You said me to remember that" + remember.read())

        elif 'voice recording' in query:
            voice_rec()

        elif 'screenshot' in query:
            screenshot()
            speak("screenshot has been taken")

        # --------------------CPU and Battery update----------------
        elif 'cpu and battery update' in query:
            cpu()
        
        # ------------------------Jokes------------------------------
        elif 'jokes' in query:
            speak("okay sir!")
            speak(pyjokes.get_joke())

        # ------------------------Camra------------------------------
        elif 'camra' in query or 'take a photo' in query:
            ec.capture(0, "your camra", "img.jpeg")
            speak("photo has been taken")

        # ----------------For background image changes---------------
        elif 'background wallpaper' in query:
            speak("Sure sir")
            ctypes.windll.user32.SystemParametersInfoW(20, 0,'E:/images/background.jpg',0)
            speak("background changed succesfully")
        
        # ------------------Username and IP Address-------------------
        elif 'host name and ip address' in query:
            hostname = socket.gethostname()
            ip_address = socket.gethostbyname(hostname)
            print(f"Hostname is: {hostname}")
            speak(f"Hostname is: {hostname}")
            print(f"IP Adderess: {ip_address}")
            speak(f"IP Adderess: {ip_address}")

        # ------------------Internate speed-------------------
        elif 'internet speed' in query:
            st = Speedtest()
            print("Your Connection's Download speed is: ", st.download())
            # speak("Your Connection's Download speed is: ", st.download())
            print("Your Connection's upload speed is: ", st.upload())
            # speak("Your Connection's upload speed is: ", st.upload())
            
        elif 'who are you' in query:
            speak("Hello Sir, I am your personal voice assistant, How can i help you")

        # elif 'internet speed' in query:
        #     speak("okay sir")
        #     os.system('cmd /k "ping localhost')

        elif 'drive information' in query:
            speak("okay sir")
            os.system('cmd /k "wmic logicaldisk list brief"')
            speak("drive information has been given")
        
        
        elif 'logout' in query:
            speak("okay sir")
            os.system("shutdown -1")

        elif 'close' in query:
            speak("Okay sir")
            os.system("shutdown /s /t 1")

        elif 'restart' in query:
            speak("Okay sir your system is restarting now")
            os.system("shutdown /r /t 1")
            
        elif 'offline' in query:
            speak("Okay sir, Have a nice day")
            quit()
            



        

    

