import os
from dotenv import load_dotenv
from openai import OpenAI

# load .env
load_dotenv()

# grab API key
api_key = os.getenv("OPENAI_API_KEY")
if not api_key:
    raise ValueError("❌ OPENAI_API_KEY not found in .env")

# connect client
client = OpenAI(api_key=api_key)

# simple test
response = client.chat.completions.create(
    model="gpt-4o-mini",
    messages=[{"role": "user", "content": "Say hello if my API key is working!"}]
)

print("✅ Connection success!")
print("Model reply:", response.choices[0].message.content)

