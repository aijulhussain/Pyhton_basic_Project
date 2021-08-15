import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import time
import webbrowser
import smtplib
import os
import win32com

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
print(voices[1].id)
engine.setProperty('voices', voices[1].id)

def speek(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour <= 12:
        speek("Good morning!")
    elif hour >= 12 and hour<=18:
        speek("Good afternoon!")
    else:
        speek("Good evening!")

    speek("I am Jarvis sir, please tell me how may I help you!")

def takeCommand():
    r = sr.Recognizer()   #it take microphone input form the user and return string output
    with sr.Microphone() as source:
        print("Spek now...")
        time.sleep(.5)
        print("Listening...")
        r.pause_threshold = 1   #seconds of non-speaking audio before a phrase is considered complete
        r.energy_threshold = 600   #minimum audio energy to consider for recording
        audio = r.listen(source)
    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print("User said: ",query)
    
    except Exception as e:
        # print(e)
        print("Say that again please...")
        return "None"
    return query

def sendEmail(to, content): # enable less secure app in gmail
    '''
    When an email client or outgoing server is submitting an email to be routed by a proper mail server, it should always use SMTP port 587 as the default port. This port, coupled with TLS encryption, will ensure that email is submitted securely and following the guidelines set out by the IETF.
    '''
    server=smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('aijulhussain14336@gmail.com', 'your-password')
    server.sendmail('aijulhussain14336@gmail.com',to, content)
    server.close()

if __name__ == "__main__":
    wishMe()
    while True:
        query = takeCommand().lower()
        '''Logic for executing task based on Query'''
        if 'wikipedia' in query:
            speek('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=3)
            speek("According to wikipedia")
            speek(results)

        elif 'open youtube' in query:
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            webbrowser.open("google.com")

        elif 'open stackoverflow' in query:
            webbrowser.open("stackoverflow.com")

        elif 'the time' in query:
            srtTime = datetime.datetime.now().strftime("%H:%M:%S")
            speek(f"Sir, the time is {strTime}")

        elif 'open vs code' in query:
            vsPath="/home/aijul/.local/share/vsliveshare/vsls-launcher"
            os.startfile(vsPath)

        elif 'open spotify' in query:
            spotifyPath = "/home/aijul/snap"
            os.startfile(spotifyPath)

        elif 'open chrome' in query:
            chromePath="/home/aijul"
            os.startfile(spotifyPath)

        elif 'email to aijul' in query:
            try:
                speek("What should I say  sir?")
                content=takeCommand()
                to = "aijulhussain14336@gmail.com"
                sendEmail(to, content)
                speek("Sir, email has been sent")
            except Exception as e:
                print(e)
                speek("Sorry sir, I am not able to send this email")
        