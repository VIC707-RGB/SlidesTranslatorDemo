import os
import time
from dotenv import load_dotenv
import google.generativeai as genai

genai.configure(api_key='PASTE YOUR API KEY HERE')

model = genai.GenerativeModel('gemini-1.5-flash')

response = model.generate_content("What is the meaning of life?")

print(response.text)