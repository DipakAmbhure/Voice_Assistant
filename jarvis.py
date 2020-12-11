import subprocess
import wolframalpha
import pyttsx3
import json
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import pyautogui
import random
import psutil
import os
import winshell
import pyjokes
import smtplib
import ctypes
import time
import requests
import shutil
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from urllib.request import urlopen


class PersonalAssistant:
    """Personal Assistant Made by GD. Able to perform following operations:

            1) Greeting according to time
            2) Tell day, date and time
            3) Play Music
            4) Play Videos
            5) Open browser
            6) Search for query on Google
            7) Send Email --Working on..
            8) small calculations 
            9) location
            10) news
            11) notes
            12) sms --working on..
            13) weather report   --working on...
            14) wikipedia
            15) lock windows, log off ,sleep,shutdown, restart
            16) changes background
            17) remember something
            18) opens codeblocks for now will add other apps also
            19) jokes
            20)
            """

    def __init__(self, name):
        self.name = name
        self.app = {"powerpoint": "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE",
                    "word": "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE",
                    "excel": "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE",
                    "cmd": "%%windir%%\\system32\\cmd.exe",
                    "codeblocks": "C:\\Program Files\\CodeBlocks\\codeblocks.exe",
                    "pycharm": "C:\\Program Files\\JetBrains\\PyCharm Community Edition 2020.2.1\\bin\\pycharm64.exe"}
        self.from_ = "ddppaagk@gmail.com"
        self.__password = "Dppaagk@123"
        self.engine = pyttsx3.init()
        voices = self.engine.getProperty('voices')
        self.engine.setProperty('voice', voices[1].id)
        self.run()


    def speak(self, text):
        """Method to convert text to speech"""
        self.engine.say(text)
        self.engine.runAndWait()

        return


    def time(self):
        """Tells info about current time"""
        time = datetime.datetime.now().strftime("%I:%M:%S")
        self.speak("the current time is")
        self.speak(time)


    def date(self):
        """Tells info about today date"""
        year = int(datetime.datetime.now().year)
        month = int(datetime.datetime.now().month)
        date = int(datetime.datetime.now().day)
        month_list = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
        self.speak("the current date is")
        self.speak(date)
        self.speak(month_list[month - 1])
        self.speak(year)

        return


    def wishme(self):
        """Wish according time"""
        #self.speak("welcome back sir!")

        #self.time()
        #self.date()

        hour = datetime.datetime.now().hour
        if 5 <= hour < 12:
            self.speak("Good morning sir!")
        elif 12 <= hour < 18:
            self.speak("Good afternoon sir!")
        elif 18 <= hour < 24:
            self.speak("Good evening sir!")
        else:
            self.speak("Good night sir!")
        self.speak(self.name + " at your service")

        return


    def usrname(self):
        """Takes username"""
        self.speak("What should i call you sir")
        self.uname = self.take_command()
        if (self.uname == "None"):
            self.usrname()
            return
        self.speak("Welcome Mister")
        self.speak(self.uname)
        columns = shutil.get_terminal_size().columns
        '''
        print("            ##########################".center(columns))
        print("Welcome Mr.", self.uname.center(columns))
        print("            ##########################".center(columns))
        '''
        self.speak("How can i Help you, Sir")

    def take_command(self):
        """Takes command using microphone"""
        r = sr.Recognizer()
        with sr.Microphone() as source:
            print("Listening.....")
            r.pause_threshold = 1
            audio = r.listen(source)
        try:
            print("Recognizing.....")
            query = r.recognize_google(audio, language="en-in")
            print("Query=", query)
        except Exception as e :
            print(e)
            self.speak("Say that again please....")
            return "None"
        return query


    def send_email(self):
        """Used to send email"""
        try:
            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.ehlo()
            server.starttls()
            server.login(self.from_, self.__password)
            self.speak("What should I say?")
            content = self.take_command()
            self.speak("whome should i send")
            to = input()
            server.sendmail(self.from_, to, content)
            server.close()
            self.speak("Email has been sent !")
        except Exception as e:
            print(e)
            self.speak("I am not able to send this email")


    def wikipedia(self, query):
        self.speak("Searching on wikipedia....")
        query = query.replace("on wikipedia", "")
        query = query.replace(self.name, "")
        try:
            result = wikipedia.summary(query, sentences=3)
            print(result)
            self.speak(result)
        except Exception as e:
            print("wikipedia exception =", e)
            self.speak("could not found page on wikipedia")
    def wolframalpha(self, query):

        app_id = "W37YJ3-P98HQR3295"
        client = wolframalpha.Client(app_id)
        indx = query.lower().split().index('calculate')
        query = query.split()[indx + 1:]
        res = client.query(' '.join(query))
        answer = next(res.results).text
        print("The answer is " + answer)
        self.speak("The answer is " + answer)


    def wolframalpha1(self, query):
        client = wolframalpha.Client("W37YJ3-P98HQR3295")
        res = client.query(query)

        try:
            print(next(res.results).text)
            self.speak(next(res.results).text)
        except StopIteration:
            print("No results")

    def music(self):
        self.speak("playing music")
        # music_dir = "D:\Music"
        try:
            music_dir = "Music"
            songs = os.listdir(music_dir)
            print(songs)
            for song in songs:
                if ".mp3" in song:
                    random = os.startfile(os.path.join(music_dir, song))
            else:
                webbrowser.open("https://www.youtube.com/watch?v=hoNb6HuNmU0")

        except:
            try:
                webbrowser.open("https://www.youtube.com/watch?v=hoNb6HuNmU0")
            except:
                self.speak("could not reach at the moment")
        exit(0)
        return


    def news(self):
        try:
            jsonObj = urlopen(
                '''https://newsapi.org / v1 / articles?source = the-times-of-india&sortBy = top&apiKey =\\times of India Api key\\''')
            data = json.load(jsonObj)
            i = 1

            self.speak('here are some top news from the times of india')
            print('''=============== TIMES OF INDIA ============''' + '\n')

            for item in data['articles']:
                print(str(i) + '. ' + item['title'] + '\n')
                print(item['description'] + '\n')
                self.speak(str(i) + '. ' + item['title'] + '\n')
                i += 1
        except Exception as e:
            print(str(e))


    def map(self, query):
        query = query.replace("where is", "")
        location = query
        self.speak("User asked to Locate")
        self.speak(location)
        webbrowser.open("https://www.google.nl / maps / place/" + location + "")


    def set_note(self):
        self.speak("What should i write, sir")
        note = self.take_command()
        file = open('jarvis.txt', 'a')
        self.speak("Sir, Should i include date and time")
        snfm = self.take_command()
        if 'yes' in snfm or 'sure' in snfm:
            strTime = datetime.datetime.now().strftime("% H:% M:% S")
            file.write(strTime)
            file.write(" :- ")
            file.write(note)
        else:
            file.write(note)
        file.close()


    def get_note(self):
        self.speak("Showing Notes")
        file = open("jarvis.txt", "r")
        print(file.read())
        self.speak(file.read(6))
        file.close()


    def update(self):
        self.speak("After downloading file please replace this file with the downloaded one")
        url = '# url after uploading file'
        r = requests.get(url, stream=True)

        with open("Voice.py", "wb") as Pypdf:

            total_length = int(r.headers.get('content-length'))

            for ch in progress.bar(r.iter_content(chunk_size=2391975),
                                   expected_size=(total_length / 1024) + 1):
                if ch:
                    Pypdf.write(ch)


    def weather(self):
        # Google Open weather website
        # to get API of Open weather
        api_key = "Api key"
        base_url = "http://api.openweathermap.org / data / 2.5 / weather?"
        self.speak(" City name ")
        print("City name : ")
        city_name = self.take_command()
        complete_url = base_url + "appid =" + api_key + "&q =" + city_name
        response = requests.get(complete_url)
        x = response.json()

        if x["cod"] != "404":
            y = x["main"]
            current_temperature = y["temp"]
            current_pressure = y["pressure"]
            current_humidiy = y["humidity"]
            z = x["weather"]
            weather_description = z[0]["description"]
            print(" Temperature (in kelvin unit) = " + str(
                current_temperature) + "\n atmospheric pressure (in hPa unit) =" + str(
                current_pressure) + "\n humidity (in percentage) = " + str(
                current_humidiy) + "\n description = " + str(weather_description))

        else:
            self.speak(" City Not Found ")


    def sms(self):
        # You need to create an account on Twilio to use this service
        account_sid = 'Account Sid key'
        auth_token = 'Auth token'
        client = Client(account_sid, auth_token)

        message = client.messages \
            .create(
            body=self.take_command(),
            from_='Sender No',
            to='Receiver No'
        )

        print(message.sid)



    def Path(self,p):
        m = ''
        for r, d, f in os.walk("C:\\"):
            for files in f:
                if files == p:
                    c = str("\\" + p)
                    a = str(os.path.join(r)) + c
                    break
        for i in range(len(a)):
            d = a[i]
            if d == '\\':
                d = '/'
            m += d
        return m
    
    def Remember_That(self):
        self.speak("What should i Remember?")
        data = self.take_command()
        self.speak("you said me to remember that " + data)
        remember = open('data.txt', 'a+')
        remember.write(data)
        remember.close()

    def tell_me(self):
        remember = open('data.txt', 'r')
        self.speak("You said me to remember that " + remember.read())

    def screen_shot(self):
        img = pyautogui.screenshot()
        a = str("D:/" + str(random.randint(1, 1000)) + ".png")
        img.save(a)
        self.speak("Done, Screenshot saved inside D Drive")

    def CPU(self):
        usage = str(psutil.cpu_percent())
        self.speak("CPU is at " + usage)
        battery = psutil.sensors_battery()
        self.speak("Battery is at ")
        self.speak(battery.percent)

    def codeblocks(self):
        a = open('Codeblocks.txt', 'w+')
        codeblockspath = a.read()
        if codeblockspath == "":
            codeblockspath = self.Path("codeblocks.exe")
            a.write(codeblockspath)
            a.close()
        for i in range(len(codeblockspath) - 1, 0, -1):
            if codeblockspath[i] == "/":
                codeblockspath = codeblockspath[0:i]
                break
        Var = os.listdir(codeblockspath)
        INDEX = Var.index("codeblocks.exe")
        os.startfile(os.path.join(codeblockspath, Var[INDEX]))
        return

    def WinPowershell(self):
        a = open('Powershell.txt', 'r+')
        powershellpath = a.read()
        if powershellpath == "":
            powershellpath = self.Path("powershell.exe")
            a.write(powershellpath)
            a.close()
        for i in range(len(powershellpath) - 1, 0, -1):
            if powershellpath[i] == "/":
                powershellpath = powershellpath[0:i]
                break
        Var = os.listdir(powershellpath)
        INDEX = Var.index("powershell.exe")
        os.startfile(os.path.join(powershellpath, Var[INDEX]))
        return 


    def run(self):
        """This method will test query for case mentioned as above."""
        # This Function will clean any
        # command before execution of this python file
        clear = lambda: os.system('cls')

        clear()
        self.wishme()
        self.usrname()

        while True:
            query = self.take_command().lower()
            if "time" in query:
                self.time()
            elif "date" in query:
                self.date()
            elif "wikipedia" in query:
                self.wikipedia(query)
            elif 'open youtube' in query:
                self.speak("opening Youtube")
                webbrowser.open("youtube.com")
            elif 'open google' in query:
                self.speak("opening Google")
                webbrowser.open("google.com")
            elif 'play music' in query or "play song" in query:
                self.music()
            elif 'mail' in query:
                self.send_email()
            elif 'how are you' in query:
                self.speak("I am fine, Thank you")
                self.speak("How are you, Sir")
            elif 'fine' in query or "good" in query:
                self.speak("It's good to know that your fine")
            elif "change my name to" in query:
                query = query.replace("change my name to", "")
                self.uname = query
            elif "change name" in query:
                self.speak("What would you like to call me, Sir ")
                self.name = self.take_command()
                self.speak("Thanks for naming me")
            elif "what's your name" in query or "what is your name" in query:
                self.speak("My friends call me")
                self.speak(self.name)
                print("My friends call me", self.name)
            elif "stop" in query or "exit" in query or "offline" in query:
                self.speak("thank you sir")
                self.speak("have a nice day")
                quit()
            elif "who made you" in query or "who created you" in query:
                self.speak("I have been created by GD.")
            elif 'joke' in query:
                self.speak(pyjokes.get_joke())
            elif "calculate" in query:
                self.wolframalpha(query)
            elif 'search' in query or 'play' in query:
                query = query.replace("search", "")
                query = query.replace("play", "")
                webbrowser.open(query)
            elif "who i am" in query:
                self.speak("If you talk then definately your human.")
            elif "why you came to world" in query:
                self.speak("I came here to to surve you as your assistant")
            elif 'is love' in query:
                self.speak("It is feeling that machiene could not have")
            elif "who are you" in query:
                self.speak("I am " + self.name + " your virtual assistant created by GD")
            elif 'reason for you' in query:
                self.speak("I was created as a Mini project by Team YZCC")
            elif 'change background' in query:
                ctypes.windll.user32.SystemParametersInfoW(20,
                                                           0,
                                                           "C:\\Users\\91866\\Pictures\\Saved Pictures",
                                                           0)
                self.speak("Background changed succesfully")
            elif 'news' in query:
                self.news()
            elif 'lock window' in query:
                self.speak("locking the device")
                ctypes.windll.user32.LockWorkStation()
            elif "log off" in query or "sign out" in query:
                self.speak("Make sure all the application are closed before sign-out")
                time.sleep(5)
                subprocess.call(["shutdown", "/l"])
            elif 'shutdown' in query:
                self.speak("Hold On a Sec ! Your system is on its way to shut down")
                subprocess.call('shutdown / p /f')
            elif "restart" in query:
                self.speak("Hold On a Sec ! Your system is on its way to restart")
                subprocess.call(["shutdown", "/r"])
            elif "hibernate" in query or "sleep" in query:
                self.speak("sleeping")
                subprocess.call("shutdown / h")
            elif 'empty recycle bin' in query or "delete recycle bin" in query:
                winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
                self.speak("Recycle Bin Recycled")
            elif "don't listen" in query or "stop listening" in query:
                self.speak("for how much second you want to stop me from listening commands")
                a = self.take_command().lower().replace("second", '')
                a = int(a)
                time.sleep(a)
                print(a)
            elif "where is" in query:
                self.map(query)
            elif "camera" in query or "photo" in query :
                ec.capture(0, self.name + " Camera ", "img.jpg")
            elif "show note" in query:
                self.get_note()
            elif "note" in query or "write" in query:
                self.set_note()
            elif "update assistant" in query:
                self.update()
            elif "weather" in query:
                self.weather()
            elif "send message " in query:
                self.sms()
            elif "wikipedia" in query:
                webbrowser.open("wikipedia.com")
            elif "Good Morning" in query:
                self.speak("A warm" + query)
                self.speak("How are you Mister")
                self.speak(self.uname)

            elif 'code blocks' in query or 'codeblocks' in query:
                self.codeblocks()

            elif 'remember' in query or 'remind me' in query:
                self.Remember_That()

            elif 'cpu' in query:
                self.CPU()
            
            elif 'screenshot' in query:
                self.screen_shot()

            elif 'powershell' in query or 'powershell' in query:
                self.WinPowershell()
            # most asked question from google Assistant
            elif "will you be my gf" in query or "will you be my bf" in query:
                self.speak("I'm not sure about, may be you should give me some time")
            elif "how are you" in query:
                self.speak("I'm fine, glad you me that")
            elif "i love you" in query:
                self.speak("It's hard to understand")
            elif "what is" in query or "who is" in query:
                self.wolframalpha1(query)
