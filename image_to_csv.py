import re
import csv
import os
from PIL import Image
import pytesseract

# If using Windows, set the path to the Tesseract executable
# pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'

def extract_emails(text):
    """Extract email addresses from text using regex"""
    email_pattern = r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+'
    return re.findall(email_pattern, text)

def main():
    # List of all image paths to process
    image_paths = [
        r'D:\coldmail\screencapture-scribd-doc-99938323-All-Companies-HR-Email-Ids-2025-07-10-23_07_04.png',
        r'D:\coldmail\screencapture-scribd-doc-99938323-All-Companies-HR-Email-Ids-2025-07-10-23_07_04-4.png',
        r'D:\coldmail\screencapture-scribd-doc-99938323-All-Companies-HR-Email-Ids-2025-07-10-23_07_04-3.png',
        r'D:\coldmail\screencapture-scribd-doc-99938323-All-Companies-HR-Email-Ids-2025-07-10-23_07_04-2.png',
        r'D:\coldmail\screencapture-scribd-doc-36540419-HR-EMail-IDs-of-Top-500-Indian-Companies-2025-07-10-15_27_14.png'
    ]
    
    all_emails = []
    
    # Check if Tesseract is available
    try:
        pytesseract.get_tesseract_version()
        print("Tesseract is available and working.")
    except Exception as e:
        print(f"Error: Tesseract not found or not working: {e}")
        print("Please make sure Tesseract is installed at C:\\Program Files\\Tesseract-OCR\\tesseract.exe")
        print("Or uncomment the tesseract_cmd line above and set the correct path.")
        return

    # Process each image
    for img_path in image_paths:
        try:
            if not os.path.exists(img_path):
                print(f"Warning: Image file not found: {img_path}")
                continue
                
            print(f"\nProcessing: {os.path.basename(img_path)}")
            
            # Load and process the image
            image = Image.open(img_path)
            text = pytesseract.image_to_string(image)
            
            # Extract emails from the text
            emails = extract_emails(text)
            print(f"Found {len(emails)} emails in {os.path.basename(img_path)}")
            
            # Add emails to the list
            for email in emails:
                all_emails.append({'image': img_path, 'email': email})
                
            # Show a preview of extracted text
            if text.strip():
                print(f"Text preview: {text[:200]}...")
            else:
                print("No text extracted from this image")
                
        except Exception as e:
            print(f"Error processing {img_path}: {e}")

    # Deduplicate and clean emails
    unique_emails = sorted(set(item['email'].lower() for item in all_emails if '@' in item['email']))
    
    print(f"\nTotal unique emails found: {len(unique_emails)}")

    # Save emails to CSV
    if unique_emails:
        with open('emails.csv', 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['email'])
            for email in unique_emails:
                writer.writerow([email])
        print(f"Successfully saved {len(unique_emails)} emails to emails.csv")
        
        # Print first few emails as preview
        print("\nFirst 10 emails found:")
        for i, email in enumerate(unique_emails[:10]):
            print(f"{i+1}. {email}")
    else:
        print("No valid emails found in the images.")

if __name__ == "__main__":
    main()