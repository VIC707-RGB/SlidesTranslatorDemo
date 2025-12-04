import google.generativeai as genai

API_KEY = 'AIzaSyDqhgiPKHCRHE1ctLwtduRbKtXQN78SYgE'
genai.configure(api_key=API_KEY)

model = genai.GenerativeModel('gemini-2.5-pro')

def chat_with_bot(prompt):
    response = model.generate_content(prompt)
    return response.text.strip()


print("Welcome to Gemini Chatbot! Type 'exit' to quit.\n")
while True:
    user = input("You: ")
    if user.strip().lower() == 'exit':
        print("Chatbot: Goodbye!")
        break
    bot = chat_with_bot(user)
    print("Chatbot:", bot)