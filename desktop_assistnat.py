import win32com.client
import os
import speech_recognition as sr
import webbrowser
import datetime
import openai
from config import apikey   #Make_sure you import your own api 


speaker = win32com.client.Dispatch("SAPI.SpVoice")

chatstr = ""

def say(text):
    speaker.Speak(text)


def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-pk")
            print(f"User said: {query}")
            return query
        except Exception: 
            return "some error occur sorry from jarvis."



def chat(query):
     
    openai.api_key = apikey
    global chatstr
    chatstr +=f"wahab: {query}\n jarvis: "

    response = openai.ChatCompletion.create(
    model="gpt-3.5-turbo",
    messages=[
            {"role": "user", "content": chatstr}
        ],
    temperature=1,
    max_tokens=256,
    top_p=1,
    frequency_penalty=0,
    presence_penalty=0
    )

    chatstr += f"{response['choices'][0]['message']['content']}\n"
    say(response["choices"][0]["message"]["content"])
    return  response["choices"][0]["message"]["content"]

        
def Ai(promt):
    
    openai.api_key = apikey

    text = f"Open ai response for pormot: {promt}\n **************************************************\n\n"

    response = openai.ChatCompletion.create(
    model="gpt-3.5-turbo",
    messages=[
            {"role": "user", "content": promt}
        ],
    temperature=1,
    max_tokens=256,
    top_p=1,
    frequency_penalty=0,
    presence_penalty=0
    )
    # print(response["choices"][0]["message"]["content"])
    try :
        text += response["choices"][0]["message"]["content"]
    except Exception:
        say("sorry sir please try again")
        print("Sorry sir please try again")

    if not os.path.exists("Openai"):
        os.mkdir("Openai")

    with open(f"Openai/{''.join(promt.split('using open AI')[1:]).strip()}.txt", 'w') as f:
        f.write(text)


    


if __name__ == "__main__":
    say("Hello sri. i am, jarvis")

    while True:
        print("Listening....")
        query = takecommand()

        sites = [["youtube","https://youtube.com"],["wikipedia","https://wikipedia.com"],["google","https://google.com"],["instagram","https://instagram.com"],["linkedin","https://www.linkedin.com"]]

        
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                say(f"opening {site[0]} sir...")
                webbrowser.open(site[1])
                exit()

        if "The Time".lower() in query.lower():
            strtime = datetime.datetime.now().strftime("%H:%M:%S")
            say(f"The time is {strtime}")

        elif "Using Open AI".lower() in query.lower():
            Ai(promt=query)

        elif 'jarvis Quit'.lower() in query.lower():
            exit()
        
        elif "Chat Reset".lower() in query.lower():
            chatstr = ""

        else:
            chat(query)
        
