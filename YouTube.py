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
else:
    print(f"API key length: {len(api_key)} characters")
    print(f"API key starts with: {api_key[:15]}...")

# Prepare headers
headers = {
    "Authorization": f"Bearer {api_key}",
    "Content-Type": "application/json",
    "HTTP-Referer": "https://localhost", # Optional. Site URL for rankings on openrouter.ai.
    "X-Title": "YouTube AI Chat", # Optional. Site title for rankings on openrouter.ai.
}

print("Making API request...")

response = requests.post(
  url="https://openrouter.ai/api/v1/chat/completions",
  headers=headers,
  data=json.dumps({
    "model": "deepseek/deepseek-chat:free",
#    "model": "google/gemini-2.0-flash-exp:free",
    "messages": [
      {
        "role": "user",
        "content": "Generate 5 short, catchy, and engaging YouTube Short title for a video about [Why Giving Investment Advice to Friends Is a Bad Idea! 🚨]. The title should be attention-grabbing, use strong language (e.g., 'Never,' 'Avoid,' 'Secret,' 'Worst,' 'Best'), and include relevant emojis for extra appeal. Target keywords related to that topic to improve discoverability. Make it compelling enough to encourage clicks while staying accurate to the content, Limit to 50 characters or less."
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
    if response.status_code == 401:
        print("🔑 Authentication Error: The API key appears to be invalid or expired.")
        print("💡 Solutions:")
        print("   1. Check if your API key is correct in the .env file")
        print("   2. Generate a new API key at: https://openrouter.ai/keys")
        print("   3. Make sure your OpenRouter account has sufficient credits")
        print("   4. Verify the API key hasn't been revoked")
    print("Response:", response.text)