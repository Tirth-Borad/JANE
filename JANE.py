import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import pywhatkit
import speedtest
from speedtest import *
import calendar
import sys
import wolframalpha
import subprocess
import tkinter
import json
import random
import operator
import smtplib
import pyjokescli
from twilio.rest import Client
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
from tabulate import tabulate
from tinytag import TinyTag

contacts = {
    "yash": "+919904366478",
    "parth": "+917972453136",
    "tirth": "+918160727649",
    "srushti": "+919409309810",
}

months = {
    "january": 1,
    "february": 2,
    "march": 3,
    "april": 4,
    "may": 5,
    "june": 6,
    "july": 7,
    "august": 8,
    "september": 9,
    "october": 10,
    "november": 11,
    "december": 12,
}

engine = pyttsx3.init("sapi5")
voices = engine.getProperty("voices")

engine.setProperty("voice", voices[1].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    speak("Hey")
    hour = int(datetime.datetime.now().hour)
    if 0 <= hour < 12:
        speak("Good Morning!")

    elif 12 <= hour < 18:
        speak("Good Afternoon!")

    else:
        speak("Good Evening!")

    speak("I am JANE. How may I Help You?")


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 0.6
        audio = r.listen(source, phrase_time_limit=3)

        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User Said : {query} \n")

        except Exception as e:
            print("Say That Again Please...")
            return "None"
        return query


def username():
    speak("What should i call you sir")
    uname = takeCommand()
    speak("Welcome Mister")
    speak(uname)
    speak("How can i help you, Sir?")


def takeMusic():
    speak("please say the song name")
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 5
        audio = r.listen(source, phrase_time_limit=3)

        try:
            print("Recognizing...")
            music = r.recognize_google(audio, language="en-in")
            print(f"User Said : {music} \n")

        except Exception as e:
            print("Say That Again Please...")
            return "None"
        return music


def takeMessage():
    speak("Speak Your Message...")
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 0.6
        audio = r.listen(source)

        try:
            print("Recognizing...")
            message = r.recognize_google(audio, language="en-in")
            print(f"User Said : {message} \n")

        except Exception as e:
            print("Say That Again Please...")
            return "None"
        return message


def takeMonth():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 0.6
        audio = r.listen(source)

        try:
            print("Recognizing...")
            month = r.recognize_google(audio, language="en-in")
            print(f"User Said : {month} \n")

        except Exception as e:
            print("Say That Again Please...")
            return "None"
        return month


def takeYear():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

        try:
            print("Recognizing...")
            Year = r.recognize_google(audio, language="en-in")
            Year = int(Year)
            print(f"User Said : {Year} \n")

        except Exception as e:
            print("Say That Again Please...")
            return "None"
        return Year


def test():
    st = speedtest.Speedtest()
    download = round(st.download())
    download = download / 1e6
    upload = round(st.upload())
    upload = upload / 1e6
    print("Your Download Speed is : ", download)
    print("Your Upload Speed is : ", upload)
    speak("f'Your Download Speed is : ' , {download}")
    speak("f'Your Upload Speed is : ' , {upload}")


def printSongs():
    onlyfiles = [
        f
        for f in os.listdir("F:\\Music\\KR$NA\\")
        if os.path.isfile(os.path.join("F:\\Music\\KR$NA\\", f))
    ]

    songsList = []
    num = 1
    for songs in onlyfiles:
        file = "F:\\Music\\KR$NA\\" + songs
        tag = TinyTag.get(file)
        song = [num, tag.title, tag.artist]
        songsList.append(song)
        num += 1

    print(
        tabulate(
            songsList,
            headers=["Sr. No.", "Title", "Artist"],
            tablefmt="grid",
            maxcolwidths=50,
        )
    )


if __name__ == "__main__":
    wishMe()
    count = 0
    while True:
        query = takeCommand().lower()

        if "wikipedia" in query:
            speak("Searching Wikipedia")
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According To Wikipedia")
            print(results)
            speak(results)

        elif "open youtube" in query:
            url = "youtube.com"
            webbrowser.register(
                "chrome",
                None,
                webbrowser.BackgroundBrowser(
                    "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
                ),
            )
            webbrowser.get("chrome").open(url)

        elif "open stack overflow" in query or "open stackoverflow" in query:
            url = "stackoverflow.com"
            webbrowser.register(
                "chrome",
                None,
                webbrowser.BackgroundBrowser(
                    "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
                ),
            )
            webbrowser.get("chrome").open(url)

        elif "open amazon" in query:
            url = "amazon.com"
            webbrowser.register(
                "chrome",
                None,
                webbrowser.BackgroundBrowser(
                    "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
                ),
            )
            webbrowser.get("chrome").open(url)

        elif "open 10fastfingers" in query:
            url = "10fastfingers.com"
            webbrowser.register(
                "chrome",
                None,
                webbrowser.BackgroundBrowser(
                    "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
                ),
            )
            webbrowser.get("chrome").open(url)

        elif "open chrome" in query:
            chromePath = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'
            os.startfile(chromePath)

        elif "open spotify" in query:
            spotifyPath = (
                'C:\\Users\\Yash\\AppData\\Local\\Microsoft\\WindowsApps\\Spotify.exe'
            )
            os.startfile(spotifyPath)

        elif "open game" in query:
            valoPath = (
                'C:\\Riot Games\\Riot Client\\RiotClientServices.exe'
            )
            os.startfile(valoPath)

        elif "open vs code" in query:
            codePath = (
                'C:\\Users\\Yash\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe'
            )
            os.startfile(codePath)

        elif "speedtest" in query or "speed test" in query:
            test()

        elif "the time" in query:
            strTime = datetime.datetime.now().strftime("%I:%M:%p")
            speak(f"Sir,The Time is {strTime}")
            print(f"The Time is {strTime}")

        elif "show calendar" in query:
            speak("Please Specify Month for Calendar")
            month = takeMonth().lower()
            mm = months[month]
            speak("Please Specify Year for Calendar")
            yy = int(takeYear())

            print(calendar.month(yy, mm))

        elif "play music" in query or "play song" in query:
            music_dir = r"F:\\Music\\KR$NA"
            songs = os.listdir(music_dir)
            # for f in range(len(songs)):
            #     print(f+1, '. ', songs[f])
            printSongs()
            try:
                music = takeMusic().lower()
                for i in songs:
                    i = i.lower()
                    if music in i:
                        song = i
                os.startfile(os.path.join(music_dir, song))
            except Exception as e:
                speak("song not found")
                print("song not found")

        elif (
            "send whatsapp message to " in query
            or "send whataspp to " in query
            or "send message to " in query
        ):
            try:
                query = query.replace("send whatsapp message to ", "")
                query = query.lower()
                Text = takeMessage()
                time = datetime.datetime.now()
                hour = int(time.strftime("%H"))
                minute = int(time.strftime("%M")) + 1
                # pywhatkit.sendwhatmsg(contacts[query],Text,hour,minute,print_wait_time=True,wait_time=50)
                pywhatkit.sendwhatmsg_instantly(contacts[query], Text, wait_time=50)
                speak("message sent")
            except Exception as e:
                print(e)
                speak("user not found")
                print("user not found")

        # QVPAVK-VR2LEKWRRV
        elif "calculate" in query:
            app_id = "QVPAVK-VR2LEKWRRV"
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index("calculate")
            query = query.split()[indx + 1 :]
            res = client.query(" ".join(query))
            answer = next(res.results).text
            print("The answer is " + answer)
            speak("The answer is " + answer)

        elif "search" in query or "play" in query:
            # app_id = "QVPAVK-VR2LEKWRRV"
            # client = wolframalpha.Client(app_id)
            # # query = query.replace("search", "")
            # # query = query.replace("play", "")
            # res = client.query(query)
            # indx = query.lower().split().index('search')
            # print(indx)
            # query = query.split()[indx + 1:]
            # res = client.query(' '.join(query))
            # print(res.results)
            # answer = next(res.results).text
            # print("The answer is " + answer)
            # speak("The answer is " + answer)

            query = query.replace("search", "")
            query = query.replace("play", "")
            url = "https://www.google.com/search?q=" + query
            webbrowser.register(
                "chrome",
                None,
                webbrowser.BackgroundBrowser(
                    "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
                ),
            )
            webbrowser.get("chrome").open(url)

        elif "news" in query:
            try:
                jsonObj = urlopen(
                    """https://newsapi.org/v2/top-headlines?country=in&sortBy=top&apiKey=30d9c2df96ef491b9f2bd99ad7f9454d"""
                )
                data = json.load(jsonObj)
                i = 1
                speak("here are some top news from india")
                print("""=============== HEADLINES FROM INDIA ============""" + "\n")
                for item in data["articles"]:
                    print(str(i) + ". " + item["title"] + "\n")
                    if item["description"] == None:
                        print(item["title"] + "\n")
                    else:
                        print(item["description"] + "\n")
                    speak(str(i) + ". " + item["title"] + "\n")
                    if i == 5:
                        break
                    i += 1
            except Exception as e:
                print(str(e))

        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.register(
                "chrome",
                None,
                webbrowser.BackgroundBrowser(
                    "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
                ),
            )
            webbrowser.get("chrome").open(
                "https://www.google.nl/maps/place/" + location + ""
            )

        elif "camera" in query or "take a photo" in query:
            ec.capture(0, "Jarvis Camera ", "img.jpg")

        elif "date" in query:
            today = datetime.date.today()
            speak(today)
            print(today)

        elif "repeat after me" in query:
            speak("What should i say?")
            print("What should i say?")
            sentence = takeCommand().lower()
            speak(sentence)
            while True:
                speak("What should i say next?")
                print("What should i say next?")
                sentence = takeCommand().lower()
                if "stop repeating" in sentence:
                    break
                speak(sentence)

        elif "shutdown" in query:
            speak("shutting down the system")
            print("shutting down the system")
            os.system("shutdown /s /t 30")

        elif "restart" in query:
            speak("restarting the system")
            print("restarting the system")
            os.system("shutdown /r /t 1")

        elif "thank you" in query:
            sys.exit()
