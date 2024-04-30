# whatsapp-automation
This Python script automates sending bulk messages via WhatsApp Web using Selenium and PyAutoGUI libraries. It reads phone numbers and messages from an Excel spreadsheet, opens the corresponding WhatsApp chats, and sends the messages. The script utilizes browser automation for WhatsApp Web and pywinauto for managing the WhatsApp Desktop app. It includes features for checking phone number formats, typing messages efficiently, and handling exceptions during message sending.

Key Features:
Automated sending of bulk WhatsApp messages
Supports sending messages with or without country codes
Utilizes PyAutoGUI for efficient typing speed
Integrates with Excel for input data management

Dependencies:
Selenium
PyAutoGUI
pywinauto
openpyxl

Usage:
Install the required dependencies.
Configure the Excel spreadsheet with phone numbers and messages.
Run the script, which will open WhatsApp chats and send messages automatically.
