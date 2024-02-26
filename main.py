# pip install speechrecognition
from typing import Any

import speech_recognition as sr
#import distutils.version

# install package pywin32 from settings to make system speak
import win32com.client

speaker = win32com.client.Dispatch("SAPI.SpVoice")

def takeCommand():
    r= sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold= 1
        audio = r.listen(source)
        try:
            query= r.recognize_google_cloud(audio,language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some error occured"


print("Enter the word")
s= input()
speaker.Speak(s)
print("Listening...")
while True:
           text= takeCommand()
           speaker.Speak(text)

 # speaker.speak


