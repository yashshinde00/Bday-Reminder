import pandas as pd
from datetime import date
from plyer import notification
import pyttsx3
import tkinter as tk

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



#Interface to show the birthday this week and this month

# Get birthdays this week and this year
this_week = []
this_month = []
for index, row in df.iterrows():
    name = row['Name']
    birthdate = row['Birthdate']
    birth_month = birthdate.month
    birth_day = birthdate.day
    if birth_month == today.month:
        if today.day <= birth_day <= today.day + 6:
            this_week.append(name)
        elif birth_day >= today.day:
            this_month.append(name)

# Create the GUI
root = tk.Tk()
root.title("Birthday Reminder")


# Birthday this week
if this_week:
    this_week_frame = tk.LabelFrame(root, text="Birthday This Week")
    this_week_frame.pack(fill="both", expand="yes", padx=10, pady=5)
    for name in this_week :
        message = f"{name} have birthday this week! "
        tk.Label(this_week_frame, text=message).pack()

#Birthday this month
if this_month:
    this_month_frame = tk.LabelFrame(root, text="Birthday This Month")
    this_month_frame.pack(fill="both", expand="yes", padx=10, pady=5)
    for name in this_month:
        message = f"{name} have birthday this month! "
        tk.Label(this_month_frame, text=message).pack()

root.mainloop()
