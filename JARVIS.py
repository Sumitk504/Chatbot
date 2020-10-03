import pyttsx3
import datetime
import speech_recognition as sr
import pyautogui
import wikipedia
import psutil
import pyjokes
import os
import ctypes
import smtplib
import webbrowser as wb
import win32com.client as winc1
from ecapture import ecapture as ec



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

# ------------------Jokes function----------------------
def jokes():
    speak(pyjokes.get_joke())



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

        elif 'date' in query:
            date()

        elif 'wikipedia' in query:
            speak("Searching...")
            query = query.replace("wikipedia","")
            result = wikipedia.summary(query, sentences = 2)
            print(result)
            speak(result)
            

        elif 'open chrome' in query:
            speak("What should I search?")
            chromepath = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'
            search = takeCommand().lower()
            wb.get(chromepath).open_new_tab(search+'.com')

        elif 'open github' in query:
            speak("Okay sir")
            githubpath = 'C:/Users/Lenovo/AppData/Local/GitHubDesktop/GitHubDesktop.exe'
            # search = takeCommand().lower()
            # wb.get(githubpathpath).open_new_tab(search)

        elif 'play some music' in query:
            songs_dir = 'D:\\Download'
            songs = os.listdir(songs_dir)
            os.startfile(os.path.join(songs_dir, songs[4]))

        elif 'remember that' in query:
            speak("Waht should i remember")
            data = takeCommand()
            speak("You said me to remember that" +data)
            remember = open('data.txt', 'w')
            remember.write(data)
            remember.close()

        elif 'do you know anything' in query:
            remember = open('data.txt', 'r')
            speak("You said me to remember that" + remember.read())

        elif 'screenshot' in query:
            screenshot()
            speak("screenshot has been taken")

        elif 'cpu' in query:
            cpu()

        elif 'joke' in query:
            jokes()

        elif 'camra' in query or 'take a photo' in query:
            ec.capture(0, "your camra", "img.jpeg")

        elif 'change the background wallpaper' in query:
            ctypes.windll.user32.SystemParametersInfoW(20, 0,'E:/images/background.jpg',0)
            speak("background changed succesfully")

        elif 'who are you' or 'tell me something about yourself' in query:
            speak("Hello Sir, I am your presonal voice assistant")

        elif 'logout' in query:
            os.system("shutdown -1")

        elif 'close' in query:
            speak("Okay sir")
            os.system("shutdown /s /t 1")

        elif 'restart' in query:
            speak("Okay sir your system is restarting now")
            os.system("shutdown /r /t 1")
            
        elif 'offline' in query:
            speak("Okay sir")
            quit()
            



        

    

