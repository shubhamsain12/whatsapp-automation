import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from pywinauto import Application
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from pyautogui import click, hotkey, press, size, typewrite
from urllib.parse import quote
import openpyxl

WIDTH, HEIGHT = size()

def check_number(number: str) -> bool:
    """Checks the Number to see if it contains the Country Code"""
    return "+" in number or "_" in number

def _web(receiver: str, message: str) -> None:
    """Opens WhatsApp Web based on the Receiver"""
    if check_number(number=receiver):
        webdriver.Chrome().get(
            "https://web.whatsapp.com/send?phone=" + receiver + "&text=" + quote(message)
        )

# Set a very short delay between characters for extremely fast typing
typewrite_delay = 0.01  # Adjust this value as needed

def send_message(message: str, receiver: str, wait_time: int) -> None:
    """Parses and Sends the Message"""
    _web(receiver=receiver, message=message)
    time.sleep(0.5)  # Wait for the WhatsApp Web page to load
    click(WIDTH / 2, HEIGHT / 2)  # Click on the message input box
    time.sleep(wait_time - 5)  # Wait for the remaining time before sending the message

    # If the receiver does not have a country code, type the message directly
    if not check_number(number=receiver):
        typewrite(message, interval=typewrite_delay)  # Type the message with a very short delay
    else:
        for char in message:
            if char == "\n":  # Check for new line characters
                hotkey("shift", "enter")  # Press Shift + Enter for a new line
            else:
                typewrite(char, interval=typewrite_delay)  # Type each character with a very short delay
    press("enter")  # Press Enter to send the message

# Load Excel file
workbook = openpyxl.load_workbook(r'C:\Users\asus\Desktop\numbers\numbers1.xlsx')
sheet = workbook.active

# Set up browser
driver = webdriver.Chrome()  # Change this according to your browser

# Message to send
message = """good morning shubham sain from Technoace."""

# Iterate through each row in Excel file
for row in sheet.iter_rows(values_only=True):
    phone_number = str(row[0])  # Assuming phone numbers are in the first column

    try:
        # Open WhatsApp link for the current phone number
        driver.get(f"https://api.whatsapp.com/send/?phone={phone_number}")

        # Wait until the "Continue to Chat" button is clickable
        continue_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="action-button"]'))
        )
        continue_button.click()

        # Wait for WhatsApp Desktop to open
        time.sleep(2)  # Adjust as needed

        # Switch to WhatsApp window using pywinauto
        app = Application().connect(title="WhatsApp", visible_only=True, timeout=10)
        window = app.window(title="WhatsApp")
        if "Desktop" in window.title:
            window.set_focus()

        # Send message to the current phone number
        send_message(message, phone_number, 10)
        print(f"Message sent successfully to {phone_number}")

    except Exception as e:
        print(f"Error sending message to {phone_number}: {e}")

    finally:
        # Wait before moving to the next number
        time.sleep(1)  # Adjust as needed

# Close the browser after sending messages to all numbers
driver.quit()
