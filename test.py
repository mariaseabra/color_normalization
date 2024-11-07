import openai

openai.api_key = "key_here"

response = openai.ChatCompletion.create(
    model="gpt-3.5-turbo",
    messages=[
        {"role": "user", "content": "Identify the color in the term 'blue'"}
    ]
)

print(response['choices'][0]['message']['content'])
