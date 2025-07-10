
# smart_emai_cold_mail
we are building smart cold malling too that will send the email using gen Ai api
=======
# Coldmail - Automated Job Application Email System

A Python-based automation tool for sending personalized job application emails to multiple companies using Outlook and AI-powered content enhancement.

## 🎯 Overview

Coldmail is designed to streamline the job application process by:
- Extracting email addresses from images using OCR (Optical Character Recognition)
- Sending personalized job application emails via Outlook
- Using AI to enhance email content for better impact
- Tracking email delivery status and maintaining logs

## 📁 Project Structure

```
coldmail/
├── image_to_csv.py          # OCR tool to extract emails from images
├── test_win32_email.py      # Basic Outlook email testing
├── test2sendmail.py         # Main email automation script
├── aitesting.py             # AI content enhancement testing
├── email.csv                # Extracted email addresses
├── email_log.csv            # Email delivery logs
├── test.csv                 # Test email list
├── vedant_singh resume.pdf  # Resume attachment
└── *.png                    # Screenshot images for OCR processing
```

## 🚀 Features

### 1. **OCR Email Extraction** (`image_to_csv.py`)
- Extracts email addresses from screenshots using Tesseract OCR
- Supports multiple image formats (PNG, JPG, etc.)
- Deduplicates and cleans email addresses
- Exports results to CSV format

### 2. **Outlook Email Automation** (`test2sendmail.py`)
- Sends emails through Microsoft Outlook using win32com
- Supports attachments (resume, cover letter)
- Configurable email templates
- Rate limiting to avoid server overload

### 3. **AI Content Enhancement** (`aitesting.py`)
- Uses OpenRouter API for AI-powered email improvement
- Enhances subject lines and body content
- Maintains professional tone and clarity

### 4. **Email Tracking & Logging**
- Logs all email delivery attempts
- Tracks success/failure status
- Timestamps for audit trail

## 🛠️ Prerequisites

### Required Software
- **Python 3.7+**
- **Microsoft Outlook** (configured with your email account)
- **Tesseract OCR** (for image processing)

### Python Dependencies
```bash
pip install pandas opencv-python pytesseract pillow pywin32 openai
```

### Tesseract Installation (Windows)
1. Download Tesseract from: https://github.com/UB-Mannheim/tesseract/wiki
2. Install to default location: `C:\Program Files\Tesseract-OCR\`
3. Add to PATH or update the path in `image_to_csv.py`

## ⚙️ Configuration

### 1. Email Settings (`test2sendmail.py`)
```python
# Configure these variables
CSV_PATH = Path("path/to/your/emails.csv")
ATTACHMENT_PATH = r"path/to/your/resume.pdf"
```

### 2. AI API Configuration
```python
client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key="your-api-key-here",
)
```

### 3. Email Template
Customize the `RAW_SUBJECT` and `RAW_BODY` variables in `test2sendmail.py` with your job application content.

## 📖 Usage

### Step 1: Extract Emails from Images
```bash
python image_to_csv.py
```
This will process all PNG images in the directory and create `emails.csv`.

### Step 2: Test Email Setup
```bash
python test_win32_email.py
```
Sends a test email to verify Outlook configuration.

### Step 3: Send Bulk Emails
```bash
python test2sendmail.py
```
Sends personalized job application emails to all addresses in the CSV file.

## 📊 Output Files

- **`emails.csv`**: Extracted email addresses from images
- **`email_log.csv`**: Detailed log of all email delivery attempts
- **Console output**: Real-time status updates during processing

## 🔧 Customization

### Email Template
Modify the `RAW_SUBJECT` and `RAW_BODY` variables in `test2sendmail.py`:
```python
RAW_SUBJECT = "Your Custom Subject Line"
RAW_BODY = """
Your personalized email content here...
"""
```

### Rate Limiting
Adjust the sleep duration between emails:
```python
time.sleep(10)  # Wait 10 seconds between emails
```

### AI Enhancement
The system automatically enhances your email content using AI. You can disable this by commenting out the `enhance_email()` function call.

## ⚠️ Important Notes

1. **Outlook Configuration**: Ensure Outlook is properly configured with your email account
2. **API Limits**: Be mindful of OpenRouter API usage limits
3. **Email Etiquette**: Respect rate limits and avoid spam filters
4. **Data Privacy**: Handle email addresses responsibly and in compliance with privacy laws

## 🐛 Troubleshooting

### Common Issues

1. **Tesseract not found**
   - Install Tesseract and update the path in `image_to_csv.py`
   - Ensure it's added to system PATH

2. **Outlook connection failed**
   - Verify Outlook is installed and configured
   - Check if Outlook is running
   - Ensure email account is properly set up

3. **API errors**
   - Verify OpenRouter API key is correct
   - Check API usage limits
   - Ensure internet connection

## 📝 License

This project is for educational and personal use. Please ensure compliance with email marketing laws and respect recipients' privacy.

## 🤝 Contributing

Feel free to submit issues and enhancement requests!

---

**Note**: This tool is designed for legitimate job applications. Please use responsibly and in accordance with email marketing best practices. 

