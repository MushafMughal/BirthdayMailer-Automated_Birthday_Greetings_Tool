# BdayMailer: Automated Birthday Greetings Tool

**BdayMailer** is a Python-based script designed to automate the sending of birthday greetings via Outlook. It reads employee information from an Excel file, checks for birthdays, and sends personalized emails with CC recipients included.

## Features

- **Automated Greeting Emails**:
  - Sends emails to employees on their birthdays.
  - Includes a customizable email body.
  - Adds CC recipients for each email.

- **Outlook Integration**:
  - Automatically dispatches emails using the Outlook COM interface.
  - Restarts Outlook if it is not running.

- **Robust Logging**:
  - Logs all operations and errors to a text file with rotation for easy debugging.

- **Timezone-Aware Scheduling**:
  - Uses Karachi (PKT) timezone to match local schedules.

## Installation

### Prerequisites

1. **Python 3.8+** installed on your system.
2. Required Python libraries:
   ```bash
   pip install pandas pytz pywin32
   ```
3. Microsoft Outlook installed and configured on your system.
4. An Excel file containing the following columns:
   - Full Name (As Per CNIC)
   - Email ID (Official)
   - Date Of Birth
   - CC1
   - CC2
### Usage
1. Place the script in your desired directory.
2. Update the following paths in the script:
   - Excel file path in pd.read_excel().
   - Log file path in log_file_path.
3. Run the script:
```bash
python bday_mailer.py
```
4. The script will monitor the time and send emails at the scheduled time (default: 00:00 PKT).

## Customization
 - Email Time: Adjust the time in the main() function:
    ```python
    if now.strftime('%H:%M') == '00:00':
    ```
   Replace 00:00 with your desired time.
 - Email Content: Modify the body variable in the send_birthday_emails() function.
## Logging
Logs are saved to the specified file (default: task.txt). It includes details of:
 - Sent emails.
 - Errors encountered.
 - Outlook restarts.
## Troubleshooting
 - Outlook Not Running: The script automatically restarts Outlook if it's not running. Ensure the path to outlook.exe is correctly set.
 - Excel File Errors: Check if the file exists and follows the expected format.
 - Timezone Issues: The script uses Karachi timezone (Asia/Karachi). Update this in the code if needed.
