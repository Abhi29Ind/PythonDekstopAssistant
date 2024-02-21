import pyttsx3 as p
import datetime as dt
import speech_recognition as sr
import os
import pywhatkit
import wikipedia
import webbrowser
from requests_html import HTMLSession
import smtplib
import time
import pyautogui

"""
ImportWarning:
1) Please Google Your User Agent and Replace the comment mention there with it.
2) For email Automation certain amendments are needed to be made in #elif "email" in query:
        a) Allow less secure apps & your Google Account
        b) in line 145 : server.login("YourEmail@mail.com","Password")
        c) in line 152 : to = "Mail Address of the recipient"
        d) in line 153 : server.sendmail("YourEmail@mail.com", to, content) 
3) Kindly import the packages according to your Operating System
"""

engine = p.init('sapi5')
voice1 = engine.getProperty('voices')
# print(voice1[0].id)
engine.setProperty('voices', voice1[0].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()


#To take voice generated input from user
def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("listening...")
        r.pause_threshold = 1
        r.adjust_for_ambient_noise(source, duration=1)
        audio = r.listen(source, timeout=5, phrase_time_limit=5)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"user said: {query}\n")
    except Exception as e:
        speak("Please tell me again...")
        return "none"
    return query


#To wish the User when program runs for the first time
def wishMe():
    hour = int(dt.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")
    elif hour>=12 and hour<16:
        speak("Good Afternoon!")
    elif hour>=16 and hour<23:
        speak("Good Evening!")
    else:
        speak("Good Night")

    speak("Hello Sir! This is Lee, how may i help you?")

if __name__=='__main__':
    wishMe()
    while True:
        query = takecommand().lower()

#To search on google
        if "google" in query:
            query = query.replace("google search","").replace("google", "").replace("search", "").replace("please search", "").replace("search on google","").replace("search in google", "").replace("on google", "").replace("in google", "")
            speak("According to Google...")
            try:
                web = "https://www.google.com/search?q=" +query
                webbrowser.open(web)
                pywhatkit.search(query)
                result = wikipedia.summary(query, sentences=2)
                speak(result)
            except Exception as e:
                speak("No Speakable output available...")

#To search or play video on Youtube
        elif "youtube" in query:
            speak("This is what I found in search!")
            query = query.replace("youtube search", "")
            query = query.replace("youtube", "")
            web = "https://www.youtube.com/results?search_query=" + query
            webbrowser.open(web)
            pywhatkit.playonyt(query)
            speak("Done, Sir")

#To open Notepad and write in Notepad
        elif "open notepad" in query:
            os.system("start notepad.exe")
            time.sleep(1)
            speak("what should i write?")
            text = takecommand().lower()
            pyautogui.typewrite(text)
            speak("Text has been written in Notepad.")
            if "bye" in text:
                os.system("taskkill /f /im notepad.exe")
            elif "close notepad" in query:
                os.system("taskkill /f /im notepad.exe")
                speak("Notepad has been closed.")

#To open Command Prompt
        elif "open command" in query:
            os.system("start cmd")

#To close the Command Prompt
        elif "close command" in query:
            os.system("taskkill /f /im cmd.exe")

#To Shutdown the Operating System
        elif "shutdown" in query or "turn off" in query:
            os.system("shutdown /s /t 1")

#To Speak Weather of any region
        elif "weather of" in query:
            s = HTMLSession()
            city = query.split("weather of")[1].strip()
            url = f'https://www.google.com/search?q=weather+of+ {city}'
            r = s.get(url, headers={'User-Agent': '# WRITE Your User Agent Here'})
            temp = r.html.find('span#wob_tm', first=True).text
            unit = r.html.find('div.vk_bk.wob-unit span.wob_t', first=True).text
            descr = r.html.find('span#wob_dc', first=True).text
            print(temp+" "+unit+" "+descr+" ")
            speak(temp + " " + unit + " " + descr + " ")

#To open MS-PowerPoint
        elif "open powerpoint" in query:
            path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE"
            os.startfile(path)
            speak("Opening power point for you sir...")

#To show the meaning of the word
        elif "meaning of" in query:
            word = query.split("meaning of")[1].strip() #To trim the sentence and take word after "meaning of"
            url = f'https://www.google.com/search?q=meaning+of+{word}'
            webbrowser.open(url)
            speak("This is what I found...")

#To Write an email and Sending it
        elif "email" in query:
            try:
                speak("Writing an email for you...")
                server = smtplib.SMTP('smtp.gmail.com', 587)
                server.ehlo()
                server.starttls()
                server.login('#mailaddress@xyz.com', '#Password')

                speak("What should I send?")
                content = takecommand().lower()

                print("Captured content:", content)

                to = "#receiver email address"
                server.sendmail("", to, content)
                speak("Sent the email....")
                server.close()
            except Exception as e:
                print(f"Error occurred while sending an email: {e}")

#To speak the time
        elif "time" in query:
            current_time = dt.datetime.now().strftime("%I:%M %p")
            speak(f"The current time is {current_time}")

#To close the voice assistant
        elif "stop" in query:
            speak("Bye Sir!")
            break