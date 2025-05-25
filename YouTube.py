import requests
import json
import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Get API key from environment variable
api_key = os.getenv('OPENROUTER_API_KEY', 'your-api-key-here')

# Debug: Check if API key is loaded
print(f"API key loaded: {'Yes' if api_key != 'your-api-key-here' else 'No'}")
if api_key == 'your-api-key-here':
    print("Warning: Using default API key. Check your .env file.")

response = requests.post(
  url="https://openrouter.ai/api/v1/chat/completions",
  headers={
    "Authorization": f"Bearer {api_key}",
    "Content-Type": "application/json",
    "HTTP-Referer": "<YOUR_SITE_URL>", # Optional. Site URL for rankings on openrouter.ai.
    "X-Title": "<YOUR_SITE_NAME>", # Optional. Site title for rankings on openrouter.ai.
  },
  data=json.dumps({
    "model": "deepseek/deepseek-chat:free",
    "messages": [
      {
        "role": "user",
        "content": "What is the meaning of life?"
      }
    ],
    
  })
)

# Check if the request was successful
if response.status_code == 200:
    # Parse the JSON response
    result = response.json()
    
    # Extract and print the AI's response
    if 'choices' in result and len(result['choices']) > 0:
        ai_response = result['choices'][0]['message']['content']
        print("AI Response:")
        print("-" * 50)
        print(ai_response)
        print("-" * 50)
    else:
        print("No response content found in the API response")
        print("Full response:", json.dumps(result, indent=2))
else:
    print(f"Error: HTTP {response.status_code}")
    print("Response:", response.text)