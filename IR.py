import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
from bs4 import BeautifulSoup
import webbrowser as web
import os
import random
import wolframalpha
from twilio.rest import Client
import requests
import subprocess
import operator
import winshell
import feedparser
import ctypes
import time
import shutil
from twilio.rest import Client
from clint.textui import progress
import win32com.client as wincl
from urllib.request import urlopen
import json
from pyowm import OWM
import re


speakEngine = pyttsx3.init("sapi5")
voices = speakEngine.getProperty('voices')
speakEngine.setProperty('voice', voices[1].id)


def IRspeak(audio):
    speakEngine.say(audio)
    speakEngine.runAndWait()


def IRcommand():
    ir = sr.Recognizer()
    # print(sr.Microphone.list_microphone_names())
    with sr.Microphone() as source:

        print("Listening...")
        ir.pause_threshold = 0.8
        ir.energy_threshold = 300
        ir.adjust_for_ambient_noise(source, duration=1)
        ir.phrase_threshold = 0.3
        audio = ir.listen(source)
        try:
            print("recognizing....")
            asked = ir.recognize_google(audio, language='en-in')
            print(f"{asked}\n")
            # asked = ir.recognize_google(audio, language='en-in')
        except Exception as e:
            IRspeak("Say it again, my love ")
            return "None"
    return asked


# def websearch():
#     path = "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe"
#     ir = sr.Recognizer()
#     # print(sr.Microphone.list_microphone_names())
#     with sr.Microphone() as source:

#         print("Listening...")
#         ir.pause_threshold = 0.8
#         ir.energy_threshold = 400
#         ir.adjust_for_ambient_noise(source, duration=0.5)
#         ir.phrase_threshold = 0.3
#         audio = ir.listen(source)
#         try:
#             print("recognizing....")
#             askedweb = ir.recognize_google(audio, language='en-in')
#             print(f"{askedweb}\n")
#             web.get(path).open(askedweb)
#             # asked = ir.recognize_google(audio, language='en-in')
#         except Exception as e:
#             IRspeak("Say it again, my love ")
#             return "None"
#     return askedweb


if __name__ == "__main__":
    IRspeak("I am harley,my love. what can i do for you ?")
    while True:
        asked = IRcommand().lower()
        # askedweb = websearch().lower()
        if "wikipedia" in asked:
            IRspeak("searching wikipedia..")
            asked = asked.replace("wikipedia", "")
            result = wikipedia.summary(asked, sentences=3)
            # lis = BeautifulSoup(result).find_all('li')
            IRspeak("According to wikipedia\n")
            print(result)
            IRspeak(result)
            IRspeak("\nThat's what i know, dear\n")

        elif "thank you" in asked:
            IRspeak("You are welcome, Mr.Joker")

        elif "you ass" in asked:
            IRspeak("sorry, i let you down . i will try hard to do my best")

        elif "well done" in asked or "good job" in asked or "nice" in asked:
            IRspeak("thank you, my dear")

        elif "love you" in asked:
            IRspeak("i know you love me, Mr.Joker")
            IRspeak("i love u too")

        elif "you there" in asked or "what are you doing" in asked:
            IRspeak("i am listenting")

        elif "your name" in asked or "about you" in asked:
            IRspeak(
                "i am isma israt jahan, aka Harley queen of Mr. joker. my version is IR 0.1. i am under development now.")
            IRspeak(
                "what can i do for you Mr. joker ?")
        elif "time" in asked:
            sTime = datetime.datetime.now().strftime("%H:%M:%S")
            IRspeak(f"the time is: {sTime}")
            print(sTime)

        elif "open youtube" in asked:
            path = "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s"
            IRspeak("opening youtube")
            web.get(path).open("youtube.com")

        elif "open github" in asked:
            path = "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s"
            IRspeak("opening github")
            web.get(path).open("github.com")

        elif "open google" in asked:
            path = "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s"
            IRspeak("opening google")
            web.get(path).open("google.com")

        elif "open facebook" in asked:
            path = "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s"
            IRspeak("opening facebook")
            web.get(path).open("facebook.com")

        elif "open linkedin" in asked:
            path = "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s"
            IRspeak("opening linkedin")
            web.get(path).open("linkedin.com")

        # elif "search" in askedweb:
        #     IRspeak("searching in google")

        elif "music" in asked:
            m_dir = "G:/musIC"
            songs = os.listdir(m_dir)
            os.startfile(os.path.join(m_dir, songs[random.randint(1, 10)]))

        elif 'weather' in asked:
            reg_ex = re.search('weather in (.*)', asked)
            if reg_ex:
                city = reg_ex.group(1)
                API_key = '017a73b4c79a346e2504ff28f8121cf4'
                o = OWM(API_key)
                owm = o.weather_manager()
                obs = owm.weather_at_place(city)
                w = owm.get_weather()
                k = w.get_status()
                x = w.get_temperature(unit='celsius')
                IRspeak('Current weather in %s is %s. The maximum temperature is %0.2f and the minimum temperature is %0.2f degree celcius' % (
                    city, k, x['temp_max'], x['temp_min']))

        elif 'lock' in asked:
            IRspeak("locking the device")
            ctypes.windll.user32.LockWorkStation()

        elif 'shutdown' in asked:
            IRspeak("shuting down the system")
            subprocess.call('shutdown /s /f')

        elif 'empty recycle bin' in asked:
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
            IRspeak("Recycle Bin is cleared")

        elif "restart" in asked:
            IRspeak(
                "Restarting the system, please make me weakup then please my love")
            subprocess.call(["shutdown", "/r"])

        elif "hibernate" in asked or "sleep" in asked:
            IRspeak("Hibernating the system")
            subprocess.call("shutdown / h")

        elif "where is" in asked:
            asked = asked.replace("where is", "")
            location = asked
            IRspeak("Location you asked is")
            IRspeak(location)
            web.open("https://www.google.nl/maps/place/" + location + "")

        elif "volume i" in asked:
            os.startfile("I:/")

        elif "django" in asked:
            os.startfile("I:/CODINGS/Web sites/Django Project")

        elif "coding" in asked:
            os.startfile("I:/CODINGS")

        elif "installation" in asked:
            os.startfile("E:/")

        elif "photoshop" in asked:
            os.startfile("E:/PhotoshopPortable")

        elif "ai" in asked:
            os.startfile("E:/IllustratorPortable")

        elif "exit" in asked or 'quit' in asked or 'bye bye' in asked or 'goodbye' in asked or 'tata' in asked or "stop" in asked or "hell" in asked or "shut up" in asked or "nothing" in asked:
            IRspeak("See you soon, love")
            exit()
