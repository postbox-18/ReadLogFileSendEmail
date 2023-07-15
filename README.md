# Read Log File in IIS Server and Send Email in SSL using Python

This repository contains a Python script that demonstrates how to read log files from an IIS server using the `watchdog` library and send SSL-encrypted email notifications using Python's `smtplib` library.

The script provides a straightforward solution for log analysis and system monitoring by monitoring log files in real-time and sending secure email notifications when log file changes occur.

## Medium Article

To learn more about the implementation details and usage of this script, please refer to the corresponding Medium article: [Python to Read IIS Log Files and Send SSL-Encrypted Emails](https://medium.com/@aravinthc18/python-to-read-iis-log-files-and-send-ssl-encrypted-emails-7316804eb7c0)

## Prerequisites

Before using this script, make sure you have the following prerequisites in place:
- An IIS server with log files enabled.
- Python installed on your system.
- Access to an SMTP server that supports SSL encryption.
- Basic knowledge of Python programming.

## Usage

1. Clone this repository to your local machine:

   ```bash
   git clone https://github.com/postbox-18/ReadLogFileSendEmail.git
   ```

2. Install the required Python dependencies:

   ```bash
   pip install -r requirements.txt
   ```

3. Configure the script:
   - Open the `config.py` file and update the configuration variables according to your setup. Provide the appropriate IIS log file path, SMTP server details, email addresses, and credentials.
   - Save the changes.

4. Run the script:
   
   ```bash
   python main.py
   ```

   The script will start monitoring the log file for changes and send email notifications when new log entries are detected.

For a step-by-step guide and detailed explanations, please refer to the Medium article mentioned above.

Feel free to customize the script to fit your specific requirements, such as adding additional log analysis logic or incorporating advanced email notification features.

For any questions or issues, please open an [issue](https://github.com/postbox-18/ReadLogFileSendEmail/issues) in this repository.

Happy log analysis and system monitoring!

Note: Make sure to refer to the Medium article for proper citations and acknowledgments.
