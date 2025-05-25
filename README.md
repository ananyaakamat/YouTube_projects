# YouTube AI Chat Project

A simple Python script that uses the OpenRouter API to interact with AI models (DeepSeek) and display responses in the terminal.

## Features

- Makes API calls to OpenRouter using the DeepSeek chat model
- Displays AI responses with formatted output in the terminal
- Includes error handling for API requests
- Clean and simple implementation

## Requirements

- Python 3.x
- requests library

## Installation

1. Clone this repository
2. Install required dependencies:
   ```bash
   pip install requests
   ```
3. Set up your API key:
   - Copy `.env.example` to `.env`
   - Add your OpenRouter API key to the `.env` file:
     ```
     OPENROUTER_API_KEY=your-actual-api-key-here
     ```
   - Or set it as an environment variable in your system

## Usage

Run the script:
```bash
python YouTube.py
```

The script will make an API call asking "What is the meaning of life?" and display the AI's response in the terminal.

## Configuration

- Update the API key in the Authorization header
- Modify the question in the "content" field to ask different questions
- Change the model if needed (currently using "deepseek/deepseek-chat:free")

## API

This project uses the OpenRouter API to access various AI models. You'll need an API key from OpenRouter to use this script.
