#Code for testing API key

import requests

api_key = ""
url = f"https://generativelanguage.googleapis.com/v1/models/gemini-2.0-flash:generateContent?key={api_key}"


data = {
    "contents": [
        {
            "role": "user",
            "parts": [{"text": "Say something inspiring"}]
        }
    ]
}

response = requests.post(url, json=data)
print(response.json())
