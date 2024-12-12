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

while True:
    now = datetime.now(pkt_timezone)
    print(f"Current time: {now.strftime('%H:%M:%S')}")  # Debugging line

    # Check if the current time is 00:00:00 PKT
    if now.strftime('%H:%M') == '00:00' and now.second == 0:
        print("Time Matched")
        today = now.strftime('%Y-%m-%d')
        print(f"Today's date: {today}")  # Debugging line

        # Filter the DataFrame for today's wishing dates
        birthday_people = bd[bd['Wishing Dates'] == today]
        print(f"Found {len(birthday_people)} people with today's wishing date.")  # Debugging line

        if birthday_people.empty:
            print("No wishing dates match today.")
        else:
            print("Checking birthdays...")  # Debugging line

            # Create an Outlook instance
            try:
                outlook = win32.Dispatch('Outlook.Application')
                log_message("Outlook instance created successfully.")
            except Exception as e:
                log_message(f"Failed to dispatch Outlook instance: {str(e)}")
                restart_outlook()  # Restart Outlook if there was an issue
                try:
                    outlook = win32.Dispatch('Outlook.Application')
                    log_message("Outlook instance created successfully after restart.")
                except Exception as e:
                    log_message(f"Failed to create Outlook instance after restart: {str(e)}")

            for index, row in birthday_people.iterrows():
                to_values = row['Email ID (Official)']
                cc_values = row['CC1'] + '; ' + row['CC2'] + ';'
                print(f"Sending email to: {to_values}, CC: {cc_values}")  # Debugging line

                # Create a new email
                mail = outlook.CreateItem(0)
                mail.To = to_values
                mail.CC = cc_values
                mail.Subject = 'Happy Birthday from Robust Support & Solutions!'

                # Construct the email body
                body = (
                    f"Hello {row['Full Name (As Per CNIC)']},\n\n"
                    "Happy Birthday from all of us at Robust Support & Solutions!\n\n"
                    "We hope your special day is filled with joy, relaxation, and your favorite activities. "
                    "Your dedication and hard work are deeply appreciated, and we are grateful to have you on our team.\n\n"
                    "Here’s to celebrating your contributions and looking forward to your continued success and happiness "
                    "in the year ahead!\n\n"
                    "Best wishes,\n\nRobust Support & Solutions Team"
                )

                mail.Body = body
                print("Email body constructed.")  # Debugging line

                # Send the email
                try:
                    mail.Send()
                    log_message(f"Email sent to: {to_values}")
                except Exception as e:
                    log_message(f"Failed to send email to {to_values}: {str(e)}")

        # Wait for 24 hours before checking again
        time.sleep(86400 - (datetime.now(pkt_timezone) - now).seconds)
    else:
        # Wait for a short time before checking the time again
        time.sleep(1)