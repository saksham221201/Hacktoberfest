import terminating
import time
import speech_recognition as sr
# importing speech recognition package from google api
import playsound  # to play saved mp3 file
from gtts import gTTS   # google text to speech
import os   # to save/open files
from selenium import webdriver  # to control browser operations


num = 1


def assistant_speaks(output):
    global num
    num +=1
    print("Assistant : " , output)
    toSpeak = gTTS(text=output, lang='en-US', slow=False)
    file = str(num)+".mp3"
    toSpeak.save(file)
    playsound.playsound(file, True)
    os.remove(file)


def get_audio():
    r = sr.Recognizer()
    audio = ''
    with sr.Microphone() as source:
        print("Speak...")
        audio = r.listen(source, phrase_time_limit=3)
    print("Stop.")
    try:
        text = r.recognize_google(audio,language='en-US')
        print("You : ", text)
        return text
    except:
        assistant_speaks("Could not understand your audio, PLease try again!")
        return 0


def search_web(input):
    driver = webdriver.Chrome()
    driver.implicitly_wait(1)
    driver.maximize_window()
    if 'youtube' in input.lower():
        assistant_speaks("Opening in youtube")
        indx = input.lower().split().index('youtube')
        query = input.split()[indx+1:]
        t="+".join(query)
        print(t)
        driver.get("http://www.youtube.com/results?search_q="+t)
        return

    elif 'wikipedia' in input.lower():

        indx = input.lower().split().index('wikipedia')
        query = input.split()[indx + 1:]
        assistant_speaks("Opening Wikipedia for "+str(query))
        driver.get("https://en.wikipedia.org/wiki/" + '_'.join(query))
        return
    elif 'music' in input:
        open_application(input)
        return
    elif 'google' in input:
            indx = input.lower().split().index('google')
            query = input.split()[indx + 1:]
            driver.get("https://www.google.com/search?q=" + '+'.join(query))
    elif 'search' in input:
        indx = input.lower().split().index('google')
        query = input.split()[indx + 1:]
        assistant_speaks("searching "+str(query))
        driver.get("https://www.google.com/search?q=" + '+'.join(query))
    else:
        driver.get("https://www.google.com/search?q=" + '+'.join(input.split()))

    return

def wait(input):
    if 'hours' in input:
        indx=input.lower().split().index('hours')
        query=(input.split()[indx-1])
        print(query)
        time.sleep(int(query)*60*60)
    elif 'seconds' in input:
        indx = input.lower().split().index('seconds')
        query = (input.split()[indx - 1])
        print(query)
        time.sleep(int(query))
    elif 'minutes' in input:
        indx = input.lower().split().index('hours')
        query = (input.split()[indx - 1])
        print(query)
        time.sleep(int(query) * 60)



def open_application(input):

    if "chrome" in input:
        assistant_speaks("Opening Google Chrome")
        os.startfile('"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"')
        return
    elif "notepad" in input:
        assistant_speaks("opening notepad")
        os.startfile("%windir%\system32\notepad.exe")
        return

    elif "firefox" in input or "mozilla" in input:
        assistant_speaks("Opening Mozilla Firefox")
        os.startfile('C:\Program Files (x86)\Mozilla Firefox\\firefox.exe')
        return
    elif " word" in input:
        assistant_speaks("Opening Microsoft Word")
        os.startfile("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Office Word 2007.lnk")
        return
    elif "excel" in input:
        assistant_speaks("Opening Microsoft Excel")
        os.startfile('C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Office Excel 2007.lnk')
        return
    else:
        assistant_speaks("Application not available")
        return
def close_application(input):
    if "chrome" in input:
        k="chrome.exe"
        terminating.terminate_task(k)
        return
    elif "notepad" in input:
        k="notepad.exe"
        terminating.terminate_task(k)
        return
    elif "word" in input:
        k="WINWORD.EXE"
        terminating.terminate_task(k)
        return
    elif "excel" in input:
        k="EXCEL.EXE"
        terminating.terminate_task(k)
        return
    elif "firefox" in input:
        k="firefox.exe"
        terminating.terminate_task(k)
        return

def process_text(input):
    try:
        if "who are you" in input or "define yourself" in input:
            speak = '''Hello, I am veronica. Your futuristic Personal Voice Assistant.
            I am here to make your life easier. 
            You can command me to perform various tasks such as calculating sums or opening applications etcetra'''
            assistant_speaks(speak)
            return
        elif "who made you" in input or "created you" in input:
            speak = "I have been created by shivansh."
            assistant_speaks(speak)
            return
        elif "crazy" in input:
            speak = """Well, there are 2 mental asylums in India."""
            assistant_speaks(speak)
            return

        elif 'open' in input:
            open_application(input.lower())
            return
        elif 'search' in input or 'play' in input or 'google' in input:
            search_web(input.lower())
            return
        elif 'wait' in input:
            wait(input.lower())
        elif 'close' in input:
            close_application(input.lower())
        else:
            assistant_speaks("I can search the web for you, Do you want to continue?")
            ans = get_audio()
            if 'yes' in str(ans) or 'yeah' in str(ans):
                search_web(input)
            else:
                return
    except Exception as e:
        print(e)
        assistant_speaks("I don't understand, I can search the web for you, Do you want to continue?")
        ans = get_audio()
        if 'yes' in str(ans) or 'yeah' in str(ans):
            search_web(input)


if __name__ == "__main__":
    name ='Saksham'
    assistant_speaks("Hey, " + name + '.')
    assistant_speaks("All Well"+'?')
    assistant_speaks("What do you want me to do?")
    assistant_speaks("I can open Applications for you ,Tell you about today's weather, google something,and can help you in certain calculations. ")

    while(1):
        assistant_speaks("Tell me what to do?")
        text = get_audio().lower()
        if text == 0:
            continue
        if "exit" in str(text) or "bye" in str(text) or "go " in str(text) or "sleep" in str(text):
            assistant_speaks("Ok bye, "+ name+'.')
            break
        process_text(text)
import wmi
def terminate_task(input):
    ti = 0
    name = input
    f = wmi.WMI()
    for process in f.Win32_Process():
        if process.name == name:
            process.Terminate()
            ti += 1
