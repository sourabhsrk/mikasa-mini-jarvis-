# -*- coding: utf-8 -*-
"""
Created on Thu Apr 25 18:29:16 2019
my api_id : ATXJVG-XLYUJQ8TYU
project: mikasa
"""
import os
import speech_recognition as sr
import win32com.client as wc
import pyttsx3
import webbrowser
import random
import wikipedia
import datetime
import sys
import wolframalpha

engine = pyttsx3.init('sapi5')

client = wolframalpha.Client('ATXJVG-XLYUJQ8TYU')

voices = engine.getProperty('voices')
engine.setProperty('voice',voices[len(voices)-1].id)

def speak(audio):
    print('Computer: ' + audio)
    engine.say(audio)
    engine.runAndWait()

def greetMe():
    currentH = int(datetime.datetime.now().hour)
    if currentH >= 0 and currentH < 12:
        speak('Good Morning!')

    if currentH >= 12 and currentH < 18:
        speak('Good Afternoon!')

    if currentH >= 18 and currentH !=0:
        speak('Good Evening!')

greetMe()

speak('Hello Sir, I am Mikasa')
speak('How may I help you?')

def internet():
    os.system("start chrome.exe")

def myCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Speak Anything :")
        audio = r.listen(source)
    try:
        user_said = r.recognize_google(audio)
        print("you said: " +user_said)
    except:
        print('Sorry sir! could not recognize your voice.try typing the command')
        user_said = str(input('command: '))

    return user_said

if __name__ == '__main__':

    while True:
        user_said = myCommand();
        user_said = user_said.lower()

        if 'open youtube' in user_said:
            speak('okay')
            webbrowser.open('www.youtube.com')

        elif 'open google' in user_said:
            speak('okay')
            webbrowser.open('www.google.co.in')

        elif 'open gmail' in user_said:
            speak('okay')
            webbrowser.open('www.gmail.com')
        
        elif(user_said == "open chrome"):
            speak('okay')
            internet()
        
        elif 'what\'s up' in user_said or 'how are you' in user_said:
            msgs = ['Just doing my thing!', 'I am fine!', 'Nice!', 'I am nice and full of energy']
            speak(random.choice(msgs))

        elif 'play music' in user_said:
            music_folder = 'D:\music\\'
            music = ['\firstsong','\secondsong']
            random_music = music_folder + random.choice(music) + '.mp4'
            os.startfile(random_music)

            speak('okay,here is your music! Enjoy!')
            
        elif 'bye' in user_said:
            speak('bye sir, have a good day.')
            sys.exit()
            
        elif(user_said == "stop program"):
            sys.exit()

        else:
            user_said = user_said
            speak('Searching...')
            try:
                try:
                    res = client.query(user_said)
                    results = next(res.results).text
                    speak('WOLFRAM-ALPHA says - ')
                    speak('Got it.')
                    speak(results)
                    
                except:
                    results = wikipedia.summary(user_said, sentences=1)
                    speak('Got it.')
                    speak('WIKIPEDIA says - ')
                    speak(results)
        
            except:
                webbrowser.open('www.google.com')
        
        speak('next command sir')
       

