from win32com.client import Dispatch
import datetime
import pytz
import speech_recognition as sr
import smtplib 



def speak(audio):
    speaking = Dispatch('SAPI.Spvoice')
    speaking.speak(audio)

def time():
    t_now = datetime.datetime.now().strftime('%H:%M:%S')
    print(t_now)
    speak("The current time is")
    speak(t_now)
def greet():
    t_hour = datetime.datetime.now().hour
    if 24>t_hour<4:
        speak("hello sir, Jarvis at your service")
    elif 12>t_hour>4:
        speak("Good Morning Sir")
    elif 18>t_hour>12:
        speak("Good afternoon sir")
    else:
        speak("Good evening sir ,i hope your day was well")

def date():
    t_date = datetime.datetime.now(tz=pytz.timezone('Asia/Kolkata'))
    speak("today's date is"+t_date.strftime('%d %B,%Y'))

def takecommand():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        print("Jarvis at your service , how may i help you")
        r.pause_threshold = 1
        command=r.listen(source)
    try:
        print("Recognizing........")
        recognised = r.recognize_google(command, language="en")
        print(recognised)
    except Exception as e :
        print(e)
        statement = "Pardon sir...."
        print(statement)
        speak(statement)
        return None
    return recognised

def getpass():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        speak("Tell your password")
        r.pause_threshold = 1
        command=r.listen(source)
    try:
        print("Recognizing........")
        recognisedpass = r.recognize_google(command, language="en")
        print(recognisedpass)
    except Exception as e :
        print(e)
        statement = "Pardon sir...."
        print(statement)
        speak(statement)
        return None
    return recognisedpass

def getemail():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        speak("Tell your email id")
        r.pause_threshold = 1
        command=r.listen(source)
    try:
        print("Recognizing........")
        gf = r.recognize_google(command, language="en")
        g=gf+"@gmail.com"
        print(g)
    except Exception as e :
        print(e)
        statement = "Pardon sir...."
        print(statement)
        speak(statement)
        return None
    return g

def getdestemail():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        speak("To whom should i send this")
        r.pause_threshold = 1
        command=r.listen(source)
    try:
        print("Recognizing........")
        gff = r.recognize_google(command, language="en")
        m=gff+"@gmail.com"
        print(m)
    except Exception as e :
        print(e)
        statement = "Pardon sir...."
        print(statement)
        speak(statement)
        return None
    return m
   

def sendemail(to,content):
    server = smtplib.SMTP("smtp.gmail.com",587)
    server.ehlo() #connecting server to gmail server
    server.starttls()#to provide security
    
    server.login(g,p)
    server.sendmail(g,to,content)
    server.close()




   

if __name__ == '__main__':
    greet()
    while True:
        query = takecommand()
        if "time" in query:
            time()
        elif "date" in query:
            date()
        
        elif "send email" in query:
            speak("Welcome to voice based email system for blind people, I am your personal assistant today")
            g=getemail()
            p=getpass()
            g="xyzsqrty@gmail.com"
            p="xyzsqr12345"
            speak("Congratulations!You are now logged in")
            try:
                speak("What should i say sir")
                content=takecommand()
                getdestemail()
                to="karmanya100@gmail.com"
                sendemail(to,content)
                speak("Email has been sent sir")
            except Exception as e:
                print(e)
                speak("Email has not been sent due to some error, please try again after sometime")
                
        elif "quit" in query:
            quit()
        
    
    