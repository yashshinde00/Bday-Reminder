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

# Get birthdays this week and this year
this_week = []
this_year = []
for index, row in df.iterrows():
    name = row['Name']
    birthdate = row['Birthdate']
    birth_month = birthdate.month
    birth_day = birthdate.day
    if birth_month == today.month:
        if today.day <= birth_day <= today.day + 6:
            this_week.append(name)
        elif birth_day >= today.day:
            this_year.append(name)
    elif birth_month > today.month:
        this_year.append(name)

# Create the GUI
root = tk.Tk()
root.title("Birthday Reminder")

# Birthday today
today_frame = tk.LabelFrame(root, text="Birthday Today")
today_frame.pack(fill="both", expand="yes", padx=10, pady=5)
for index, row in df.iterrows():
    name = row['Name']
    birthdate = row['Birthdate']
    birth_month = birthdate.month
    birth_day = birthdate.day
    if birth_month == today.month and birth_day == today.day:
        message = f"{name} is celebrating their birthday today!"
        tk.Label(today_frame, text=message).pack()

# Birthday this week
if this_week:
    this_week_frame = tk.LabelFrame(root, text="Birthday This Week")
    this_week_frame.pack(fill="both", expand="yes", padx=10, pady=5)
    for name in this_week:
        message = f"{name} is celebrating their birthday this week!"
        tk.Label(this_week_frame, text=message).pack()

# Birthday this year
if this_year:
    this_year_frame = tk.LabelFrame(root, text="Birthday This Year")
    this_year_frame.pack(fill="both", expand="yes", padx=10, pady=5)
    for name in this_year:
        message = f"{name} will celebrate their birthday this year!"
        tk.Label(this_year_frame, text=message).pack()

root.mainloop()
