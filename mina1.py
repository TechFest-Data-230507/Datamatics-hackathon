import imaplib
import email
from email.header import decode_header
import time
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
import google.generativeai as genai
import requests

analyzer = SentimentIntensityAnalyzer()

genai.configure(api_key="yourapikey")  # Replace with your actual API key
model = genai.GenerativeModel('gemini-1.5-flash')
# Email account credentials
EMAIL = "target-mail-id"
PASSWORD = "password"  # Consider using an app password for security
IMAP_SERVER = "imap.gmail.com"

previous_email_ids = set()

def connect_to_email():
    """Connect to the email server and log in."""
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL, PASSWORD)
    return mail

def fetch_emails(mail):
    """Fetch new emails and display them."""
    global previous_email_ids
    # Select the mailbox you want to check (e.g., "INBOX")
    mail.select("INBOX")

    # Search for all emails in the inbox
    status, messages = mail.search(None, "ALL")

    # Convert the result to a set of email IDs
    email_ids = set(messages[0].split())

    # Debugging: print the currently fetched email IDs
    

    # Determine which emails are new
    new_email_ids = email_ids - previous_email_ids
    
    previous_email_ids = email_ids  # Update the previous email IDs

    # Store emails in a list for processing
    new_emails = []

    # Loop through all the new email IDs
    for email_id in sorted(new_email_ids, reverse=True):  # Ensure IDs are sorted in descending order
        # Fetch the email by ID
        status, msg_data = mail.fetch(email_id, "(RFC822)")

        for response_part in msg_data:
            if isinstance(response_part, tuple):
                # Parse a bytes email into a message object
                msg = email.message_from_bytes(response_part[1])

                # Decode the email subject
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding if encoding else "utf-8")

                # Get the sender email
                from_ = msg.get("From")

                # Get the email date
                date_ = msg.get("Date")
                email_date = email.utils.parsedate_to_datetime(date_)

                # Get the email body
                body = ""
                if msg.is_multipart():
                    for part in msg.walk():
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        if "attachment" not in content_disposition:
                            if content_type == "text/plain":
                                body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                                break
                else:
                    body = msg.get_payload(decode=True).decode("utf-8", errors="ignore")

                # Store the email data in a dict
                new_emails.append({
                    "from": from_,
                    "subject": subject,
                    "date": email_date,
                    "body": body
                })

    # Sort emails by date descending
    new_emails.sort(key=lambda x: x["date"], reverse=True)
    #return new_emails
    if new_emails:
        print("\n--- New Emails Fetched ---\n")
        for mail_info in new_emails:
            formatted_date = mail_info["date"].strftime("%Y-%m-%d %H:%M:%S")
            print(f"Sender: {mail_info['from']}")
            print(f"Subject: {mail_info['subject']}")
            print(f"Date: {formatted_date}")
            print(f"Body: {mail_info['body'][:100]}...")  # Print first 100 characters of the body
            print("-" * 50)

            sentiment_scores = analyzer.polarity_scores(mail_info['body'])
            compound_score=sentiment_scores['compound']
            summary=generate_summary(mail_info['body'])
            submit_to_gform(mail_info['body'],getsentiment(compound_score),compound_score,summary[0],summary[1],mail_info['from'])
            time.sleep(5)


def datastuff(themail):
    thedate=[themail['from'],themail['subject'],themail['body']]


def submit_to_gform(suggestion,suggestionres,score,category,orderid,userid):

    # Google Form URL (replace with your actual form URL)
    form_url = 'https://docs.google.com/forms/d/1ubB3TjD-29z8nxjFxkcUe_RTAiQN3AUjrl6OoahzezI/formResponse'  # Replace with actual form URL
    
    # Map the data to form fields
    form_data = {
        'entry.1877889307': suggestion,  # suggestion/feedback
        'entry.2033190925': suggestionres, #positive/negative/neutral (sentiment result)
        'entry.1341410943': score, #sentiment scores
        'entry.1329749283' : 'Service',
        'entry.1817253318': orderid,
        'entry.689965723': userid
        
    }

    # Submit the data
    response = requests.post(form_url, data=form_data)
    
    # Check the response
    if response.status_code == 200:
        print("Form submission successful!")
    else:
        print(f"Form submission failed with status code {response.status_code}")

def generate_summary(mail):
    
    
    # Generate content using the Generative AI model
   
    response1=model.generate_content("please categorise the given feedback into either product or service or delivery feedback and produce one word answer:\n"+mail)
    response2=model.generate_content("please find the order id in the following text. if it exists output only the orderid, else output \"no\" :\n"+mail)
    return [response1.text,response2.text]

def getsentiment(compound_score):
    if compound_score>=0.05:
        return "Positive"
    elif compound_score>-0.05 and compound_score<0.05:
        return "Neutral"
    elif compound_score<=-0.05:
        return "Negative"

connect_to_email()

def main():
    """Main function to run the email checker."""
    mail = connect_to_email()
    print('conncted')
    try:
        while True:
            fetch_emails(mail)
                #print(datastuff(i))
            time.sleep(10)  # Wait for 10 seconds before checking again
    except KeyboardInterrupt:
        print("Email checking stopped.")
    finally:
        mail.logout()
main()
