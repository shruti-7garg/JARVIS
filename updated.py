import random
import pyttsx3
import win32com.client
import datetime
import wikipedia
import webbrowser
import os
import smtplib
import speech_recognition as sr

speaker = win32com.client.Dispatch("SAPI.SpVoice")

#engine= pyttsx3.init('sapi5')
#voices = engine.getProperty('voices')

#def speak(audio):
 #   engine.say(audio)
  #  engine.runAndWait()

#if __name__=="__main__":
 #   speak("harry is a good boy")
  #  print("hello")


# SPEAK FUNCTION
def speak():
    print("enter the word")
    s = input()
    speaker.Speak(s)


# WISH FUNCTION
def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speaker.Speak("Good Morning!")
    elif hour>12 and hour<18:
        speaker.Speak("Good Afternoon!")
    else:
        speaker.Speak("Good Evening!")

    speaker.Speak("I am Jarvis ,how can I help you")


# SPEECH TO TEXT FUNCTION USING GOOGLE
def takeCommand():
    # it takes microphone input from the user and returns string output
    r= sr.Recognizer()
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source)
        print("Listening...")
        #r.pause_threshold = 1
        audio = r.listen(source)
    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")
    except Exception as e:
        #print(e)
        print("say that again please...")
        query= "None"
    return query


# SEND EMAIL FUNCTION
def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com',587)
    server.ehlo()
    server.starttls()
    server.login('sg456560@gmail.com', 'lgcwohacvejsbqmu')
    server.sendmail('sg456560@gmail.com',to,content)
    server.close()

# MAIN FUNCTION
#speak()
wishMe()
while True:
    query = takeCommand().lower()

 # logic for executing tasks based on query
    if 'wikipedia' in query:
        speaker.Speak("Searching wikipedia...")
        query = query.replace("wikipedia","")
        results = wikipedia.summary(query, sentences=2)
        speaker.Speak("According to wikipedia")
        print(results)
        speaker.Speak(results)

    elif 'open youtube' in query:
        webbrowser.open("youtube.com")

    elif 'open google' in query:
        webbrowser.open("google.com")

    elif 'open striver sheet' in query:
        webbrowser.open("takeUforward.org")

    elif 'play music' in query:
         music_dir = "D:\\front-end projects\\Spotify Clone\\songs"
         songs= os.listdir(music_dir)
         print(songs)
         number = random.randint(0,6)
         os.startfile(os.path.join(music_dir, songs[number]))

    elif 'the time' in query:
        strTime = datetime.datetime.now().strftime("%H:%M:%S")
        speaker.Speak(f"Sir, the time is {strTime}")

    elif 'open code' in query:
        codePath ="C:\\Users\\hp\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
        os.startfile(codePath)

# import contacts ans save in dictionary - pending
    # admin console - pending
    elif 'email to dhruv' in query:
         try:
             speaker.Speak("What should I say?")
             content =takeCommand()
             to= "dhruvgarg931@gmail.com"
             sendEmail(to, content)
             speaker.Speak("Email has been sent")
         except Exception as e:
             print(e)
             speaker.Speak("Sorry my friend Shruti, I am not able to send this email")



