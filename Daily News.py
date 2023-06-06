from win32com.client import Dispatch
import requests
import json
import re
# f1c5f32d716649c2a84ee29003cd7616


def speak(string):
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(string)


r = requests.get("https://newsapi.org/v2/top-headlines?country=in&apiKey=f1c5f32d716649c2a84ee29003cd7616")
parse = json.loads(r.text)
News = parse["articles"]

for i in News:
    dict1 = i
    # content1 = re.sub(r'\[\+\d+\s?chars\]', '', dict1['title'])
    # content = re.sub(r'\bInst\b', '', content1, flags=re.MULTILINE)
    # print(dict1['title'], '\n', dict1['content'], '\n')
    print(dict1['title'])
    speak(dict1['title'])

