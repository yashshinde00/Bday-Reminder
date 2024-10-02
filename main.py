import pandas as pd
from datetime import date
from plyer import notification
import pyttsx3

# Initialize pyttsx3 engine
engine = pyttsx3.init()
engine.setProperty("rate", 120)  # Set the speaking rate (words per minute)
voices = engine.getProperty("voices")
engine.setProperty("voice", voices[1].id)  # Select the voice (index 0 is for English)

# Load the Excel file with the birthdays
df = pd.read_excel('birthdays.xlsx')

# Get today's date
today = date.today()

# Loop through each row in the Excel file
for index, row in df.iterrows():

    # Get the person's name and birthdate
    name = row['Name']
    birthdate = row['Birthdate']

    # Extract the month and day from the birthdate
    birth_month = birthdate.month
    birth_day = birthdate.day
    
    # Check if the person's birthday is today
    if birth_month == today.month and birth_day == today.day:

        # Print and speak the birthday message
        message = f"Today is {name}'s birthday!"
        print(message)
        engine.say(message)
        engine.runAndWait()

        # Show a notification
        notification.notify(
            title="Birthday Reminder",
            message=message,
            timeout=5,
        
         )

