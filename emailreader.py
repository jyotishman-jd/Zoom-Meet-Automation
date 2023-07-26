import imaplib
import email
import yaml
import re
import smtplib
from email.mime.text import MIMEText
from meetauto import schedule_zoom_meeting

# Function to extract the  details from the email body
def extract_meeting_details(body):
    # Create the regular expressions
    date_pattern = re.compile(r'date\s*:\s*(.+)', re.IGNORECASE)
    time_pattern = re.compile(r'time\s*:\s*(\d{1,2}:\d{2})\s*([APMapm]+)', re.IGNORECASE)
    duration_pattern = re.compile(r'duration\s*:\s*(\d+)', re.IGNORECASE)
    topic_pattern = re.compile(r'topic\s*:\s*(.+)', re.IGNORECASE)

    # Match the Patterns
    date_match = date_pattern.search(body)
    time_match = time_pattern.search(body)
    duration_match = duration_pattern.search(body)
    topic_match = topic_pattern.search(body)
    
    #Insert the extracted Data into variables
    date = date_match.group(1) if date_match else None
    time = time_match.group(1) if time_match else None
    time_period = time_match.group(2) if time_match else None
    duration = duration_match.group(1) if duration_match else None
    topic = topic_match.group(1) if topic_match else None

    return date, time, time_period , duration, topic


# Loading the gmail credentials
with open("credentials.yml") as f:
    credentials = yaml.load(f, Loader=yaml.FullLoader)

# Importing from Credentials
username = credentials["username"]
password = credentials["password"]

# IMAP connection details
imap_url = 'imap.gmail.com'

mail = imaplib.IMAP4_SSL(imap_url)

# Log in part
mail.login(username, password)

# Select the Inbox to fetch messages
mail.select('Inbox')

search_key = 'SUBJECT'
search_value = 'zoom meeting'

status, data = mail.search(None, f'{search_key} "{search_value}"')

#Check for a matching Mail

if status == 'OK':
    email_ids = data[0].split()
    email_count = len(email_ids)
    #print(f"Found {email_count} email(s) with subject '{search_value}'.")
    print("We have a matching Mail")

    if email_count > 0:
        latest_email_id = email_ids[-1]  # Get the latest email id
        status, email_data = mail.fetch(latest_email_id, '(RFC822)')

        if status == 'OK':
            raw_email = email_data[0][1]
            message = email.message_from_bytes(raw_email)

            print("_________________________________________")
            print("Subject:", message['Subject'])
            print("From:", message['From'])
            print("Body:")
            
            # Fetch subject and sender
            subject = message['Subject']
            sender = message['From']
            # Extract the text body of the email
            if message.is_multipart():
                for part in message.walk():
                    if part.get_content_type() == 'text/plain':
                        body = part.get_payload()
                        print(body)
                        # Extract meeting details from the email body
                        date, time, time_period ,duration, topic = extract_meeting_details(body)
                        print("Date:", date)
                        print("Time:", time)
                        print("Duration:", duration)
                        print("Topic:", topic)
            else:
                body = message.get_payload()
                print(body)
                # Extract meeting details from the email body
                date, time, duration, topic = extract_meeting_details(body)
                
            LinkGenerator = schedule_zoom_meeting(topic,date,time,time_period,sender,duration)
            
            
            # Send a reply email
            reply_body = "This is an automated reply to your email.\n" + "\n".join(LinkGenerator)
            reply_subject = f"Re: {subject}"
            reply_email = username  # Admin Email Replacable
            
            msg = MIMEText(reply_body)
            msg['Subject'] = reply_subject
            msg['From'] = reply_email
            msg['To'] = sender

            try:
                smtp_server = 'smtp.gmail.com'  # Replacable 
                smtp_port = 587  # Replaceable

                smtp = smtplib.SMTP(smtp_server, smtp_port)
                smtp.starttls()
                smtp.login(username,password)  # Replacable
                smtp.sendmail(reply_email, sender, msg.as_string())
                smtp.quit()

                print("Reply email sent successfully!")
            except Exception as e:
                print("Failed to send reply email:", e)
   
    else:
        print("No emails found with subject '{search_value}'.")
else:
    print("Failed to fetch emails from the inbox.")

mail.logout()
