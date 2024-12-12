import win32com.client as win32
import pandas as pd
from datetime import datetime
import pytz
import time
import os
import logging
from logging.handlers import RotatingFileHandler

# Set up logging with rotation
log_file_path = "task.txt"
handler = RotatingFileHandler(log_file_path, maxBytes=5_000_000, backupCount=5)
logging.basicConfig(level=logging.INFO, handlers=[handler], format='%(asctime)s - %(levelname)s - %(message)s')

logging.info("Script started.")

# Load the Excel file and preprocess the data
try:
    bd = pd.read_excel('Birthday Months & Emails.xlsx')
    bd = bd[['Full Name (As Per CNIC)', 'Email ID (Official)', 'Date Of Birth', 'CC1', 'CC2']]
    bd['Date Of Birth'] = pd.to_datetime(bd['Date Of Birth'])
except Exception as e:
    logging.error(f"Error loading or processing the Excel file: {e}")
    raise

# Prepare wishing dates
pkt_timezone = pytz.timezone('Asia/Karachi')
current_year = datetime.now(pkt_timezone).year
bd['Wishing Dates'] = bd['Date Of Birth'].apply(lambda x: x.replace(year=current_year))

def restart_outlook():
    """Restart Outlook to ensure it's running."""
    try:
        os.system('taskkill /IM outlook.exe /F')
        logging.info("Outlook process terminated.")
        time.sleep(5)
        os.startfile("outlook.exe")
        logging.info("Outlook restarted.")
        time.sleep(10)
    except Exception as e:
        logging.error(f"Failed to restart Outlook: {e}")

def send_birthday_emails():
    """Send birthday emails to employees."""
    today = datetime.now(pkt_timezone).strftime('%Y-%m-%d')
    birthday_people = bd[bd['Wishing Dates'] == today]

    if birthday_people.empty:
        logging.info("No birthdays to celebrate today.")
        return

    try:
        outlook = win32.Dispatch('Outlook.Application')
        logging.info("Outlook instance created successfully.")
    except Exception as e:
        logging.error(f"Failed to create Outlook instance: {e}")
        restart_outlook()
        try:
            outlook = win32.Dispatch('Outlook.Application')
            logging.info("Outlook instance created successfully after restart.")
        except Exception as e:
            logging.error(f"Failed to create Outlook instance after restart: {e}")
            return

    for _, row in birthday_people.iterrows():
        to_values = row['Email ID (Official)']
        cc_values = f"{row['CC1']}; {row['CC2']}"
        subject = "Happy Birthday from Robust Support & Solutions!"
        body = (
            f"Hello {row['Full Name (As Per CNIC)']},\n\n"
            "Happy Birthday from all of us at Robust Support & Solutions!\n\n"
            "We hope your special day is filled with joy, relaxation, and your favorite activities. "
            "Your dedication and hard work are deeply appreciated, and we are grateful to have you on our team.\n\n"
            "Here’s to celebrating your contributions and looking forward to your continued success and happiness "
            "in the year ahead!\n\n"
            "Best wishes,\n\nRobust Support & Solutions Team"
        )

        try:
            mail = outlook.CreateItem(0)
            mail.To = to_values
            mail.CC = cc_values
            mail.Subject = subject
            mail.Body = body
            mail.Send()
            logging.info(f"Email sent to {to_values}")
        except Exception as e:
            logging.error(f"Failed to send email to {to_values}: {e}")

def main():
    """Main loop to monitor and send emails at the scheduled time."""
    while True:
        now = datetime.now(pkt_timezone)
<<<<<<< HEAD
        if now.strftime('%H:%M') == '00:00':  # Adjust time as needed
=======
        if now.strftime('%H:%M') == '04:24':  # Adjust time as needed
>>>>>>> 8eef053b78d2e440d879e497bff0d7a4204c09fa
            send_birthday_emails()
            time.sleep(86400)  # Wait 24 hours
        else:
            time.sleep(1)

if __name__ == "__main__":
    main()