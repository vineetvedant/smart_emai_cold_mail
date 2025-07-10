from openai import OpenAI

client = OpenAI(
  base_url="https://openrouter.ai/api/v1",
  api_key="put your api key here"
)

completion = client.chat.completions.create(
  extra_headers={
    "HTTP-Referer": "<YOUR_SITE_URL>",  # Optional
    "X-Title": "<YOUR_SITE_NAME>",      # Optional
  },
  model="mistralai/mistral-7b-instruct:free",
  messages=[
    {
      "role": "user",
      "content": "temperature of the delhi"
    }
  ]
)
print(completion.choices[0].message.content)
