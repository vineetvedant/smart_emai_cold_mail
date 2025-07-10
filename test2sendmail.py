import win32com.client as win32
from pathlib import Path
import pandas as pd
from datetime import datetime
from openai import OpenAI
import time

# ====== CONFIGURE THESE ======
CSV_PATH = Path("D:\coldmail\test.csv")  # CSV file with a column 'email'
LOG_PATH = "email_log.csv"
ATTACHMENT_PATH = r"D:\coldmail\vedant_singh resume.pdf"

# OpenAI/OpenRouter API config
client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key="put your api key here"
)

RAW_SUBJECT = "Python developer || Vedant Singh"
RAW_BODY = """
I am excited to apply for a role in Python Development, Data Analysis, or Data Science at your company. I have a strong background in software development, computer vision, and AI, with hands-on experience in building APIs and working on real-world projects.

During my internship at Proeffico Solutions, I worked on:
- Deploying AI models (YOLOv8) for employee monitoring and object detection.
- Building C++ applications for IoT systems.
- Integrating voice and text automation using generative AI and RAG pipelines.

At Tata Motors, I contributed to real-time safety systems using Python and object detection tools like YOLO.

Some of my key personal projects include:
- AI Ticketing System for the Visually Impaired – Developed with Python and Braille output.
- Object Tracking System – Built using Arduino and sensors.
- REST API Development – Created APIs for data and model deployment.
- Text-to-Audio Converter – Turned PDF content into audiobooks using Python.
- Parallel Programming – Optimized C++ code using OpenMP.

I am skilled in Python, C++, OpenCV, TensorFlow, PyTorch, and cloud tools, and I enjoy solving real-world problems through technology.

I’d love the opportunity to contribute to Company’s projects and be a part of your innovative team. Thank you for considering my application.

Best regards,
Vedant Singh
+91-8873001000
LinkedIn - https://linkedin.com/in/vedant-singh-2550b2202
GitHub - https://github.com/vineetvedant
"""

def enhance_email(subject, body):
    prompt = (
        f"Improve the following job application email for clarity, professionalism, and impact. "
        f"Return a subject line and a body in markdown format. "
        f"Subject: {subject}\n\nBody:\n{body}"
    )
    completion = client.chat.completions.create(
        model="mistralai/mistral-7b-instruct:free",
        messages=[{"role": "user", "content": prompt}]
    )
    content = completion.choices[0].message.content
    # Attempt to parse subject and body from the response
    lines = content.splitlines()
    new_subject = subject
    new_body = body
    for i, line in enumerate(lines):
        if line.lower().startswith("subject:"):
            new_subject = line.split(":", 1)[1].strip()
        if line.lower().startswith("body:"):
            new_body = "\n".join(lines[i+1:]).strip()
            break
    return new_subject, new_body

def send_email(to_email, subject, body, attachment_path=None):
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = to_email
        mail.Subject = subject
        mail.Body = body

        if attachment_path and Path(attachment_path).exists():
            mail.Attachments.Add(str(Path(attachment_path).resolve()))

        mail.Send()
        return "Sent"
    except Exception as e:
        return f"Failed: {e}"

def main():
    # Try reading with header first
    try:
        df = pd.read_csv(str(CSV_PATH))
        if 'email' not in df.columns:
            # If no 'email' column, try reading without header and assign column name
            df = pd.read_csv(str(CSV_PATH), header=None, names=['email'])
    except Exception as e:
        print(f"Failed to read CSV: {e}")
        return

    # Remove empty rows and strip whitespace
    df['email'] = df['email'].astype(str).str.strip()
    df = df[df['email'].str.contains('@') & (df['email'] != '')]

    if df.empty:
        print("No valid emails found in CSV. Please check your file.")
        return

    # Enhance email with OpenRouter/OpenAI
    print("Enhancing email content with AI...")
    enhanced_subject, enhanced_body = enhance_email(RAW_SUBJECT, RAW_BODY)
    print("Enhanced subject:", enhanced_subject)
    print("Enhanced body (first 200 chars):", enhanced_body[:200], "...")

    # Prepare log
    log_entries = []

    for i, (idx, row) in enumerate(df.iterrows()):
        to_email = row['email']
        status = send_email(to_email, enhanced_subject, enhanced_body, ATTACHMENT_PATH)
        print(f"{to_email}: {status}")
        log_entries.append({
            "timestamp": datetime.now().isoformat(),
            "email": to_email,
            "status": status
        })
        
        # Sleep for 10 seconds between emails to avoid overwhelming the server
        if i < len(df) - 1:  # Don't sleep after the last email
            print("Waiting 10 seconds before sending next email...")
            time.sleep(10)

    # Write log
    log_df = pd.DataFrame(log_entries)
    log_df.to_csv(LOG_PATH, index=False)
    print(f"Log written to {LOG_PATH}")

if __name__ == "__main__":
    main()
