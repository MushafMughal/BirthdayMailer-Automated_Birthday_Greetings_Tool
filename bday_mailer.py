import win32com.client as win32
import pandas as pd
from datetime import datetime
import pytz
import time
import os


# Function to log messages to a file
def log_message(message):
    with open("task.txt", 'a') as file:
        file.write(f'{datetime.now()}: {message}\n')

log_message("The script started running.")


# Load the Excel file and select relevant columns
bd = pd.read_excel('Birthday Months & Emails.xlsx')
bd = bd[['Full Name (As Per CNIC)', 'Email ID (Official)', 'Date Of Birth', 'CC1', 'CC2']]

# Convert 'Date Of Birth' to datetime
bd['Date Of Birth'] = pd.to_datetime(bd['Date Of Birth'])

# Create 'Wishing Dates' by updating the year to the current year in Karachi timezone
pkt_timezone = pytz.timezone('Asia/Karachi')
current_year = datetime.now(pkt_timezone).year
bd['Wishing Dates'] = bd['Date Of Birth'].apply(lambda x: x.replace(year=current_year))



def restart_outlook():
    # Kill any existing Outlook processes
    os.system('taskkill /IM outlook.exe /F')
    log_message("Existing Outlook process killed.")
    
    # Wait a moment to ensure the process has fully terminated
    time.sleep(5)

    # Attempt to launch Outlook directly before using COM object
    os.startfile("outlook.exe")
    log_message("Outlook restarted.")
    
    # Give Outlook a moment to initialize
    time.sleep(10)

while True:
    now = datetime.now(pkt_timezone)
    print(f"Current time: {now.strftime('%H:%M:%S')}")  # Debugging line

    # Check if the current time is 04:24:00 PKT
    if now.strftime('%H:%M') == '21:44' and now.second == 0:
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