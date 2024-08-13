Explain:          
WhatsApp Automated Messaging Script
This Python script automates the process of sending personalized WhatsApp messages to clients using data from an Excel spreadsheet. The script loads client data, including names, phone numbers, and special dates (like birthdays), and sends a customized message via WhatsApp Web.

How It Works
Dependencies:

openpyxl: For reading data from the Excel spreadsheet.
urllib.parse: For encoding the message text to be URL-safe.
webbrowser: To open WhatsApp Web with the pre-filled message.
time.sleep: To add delays between actions.
pyautogui: For automating the clicks on the WhatsApp Web interface.
Setup:

Ensure you have a file named clientes.xlsx with a worksheet named Planilha1.
The worksheet should contain at least three columns:
Column A: Client Name
Column B: Client Phone Number (in the format recognized by WhatsApp)
Column C: Special Date (e.g., Birthday)
How the Script Works:

The script starts by loading the clientes.xlsx file and reading the client data from Planilha1.
For each client, it constructs a personalized message, including the client's name and special date.
The script then encodes the message for URL use and opens WhatsApp Web to send the message.
The script uses pyautogui to locate the "Send" button on the screen and clicks it to send the message.
If any errors occur (e.g., if the "Send" button is not found), the client's information is logged into an erros.csv file.
Execution:

The script includes sleep delays to allow for the loading time of WhatsApp Web and the detection of the "Send" button.
Once a message is sent, the script closes the WhatsApp Web tab and proceeds to the next client.
Error Handling:

If the script is unable to send a message to a client, it prints an error message and logs the client's phone number and name in the erros.csv file for later review.
Prerequisites
Python 3.x
The required Python libraries (openpyxl, urllib, pyautogui, time)
An Excel file named clientes.xlsx with the correct data structure
Notes
Make sure the enviar.png image is available in the script directory. This image should be a screenshot of the "Send" button on WhatsApp Web.
The script assumes that WhatsApp Web is already logged in on the default browser
