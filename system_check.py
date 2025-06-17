import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Test API key
api_key = os.getenv('OPENROUTER_API_KEY', 'not-found')
print(f"🔑 API Key status: {'✅ Found' if api_key != 'not-found' and api_key != 'your-api-key-here' else '❌ Not found'}")

# Test Excel file
EXCEL_FILE_PATH = r"D:\Anant\Youtube\ValueProITGyan\YouTubeVideosList.xlsx"
print(f"📁 Excel file exists: {'✅ Yes' if os.path.exists(EXCEL_FILE_PATH) else '❌ No'}")

# Test imports
try:
    import requests
    print("✅ requests library imported")
except ImportError:
    print("❌ requests library not found")

try:
    import openpyxl
    print("✅ openpyxl library imported")
except ImportError:
    print("❌ openpyxl library not found")

try:
    from docx import Document
    print("✅ python-docx library imported")
except ImportError:
    print("❌ python-docx library not found")

print("\n🎯 System ready for automation!")
