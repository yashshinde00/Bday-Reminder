import pandas as pd
from datetime import date
import tkinter as tk

# Load the Excel file with the birthdays
df = pd.read_excel('birthdays.xlsx')

# Get today's date
today = date.today()

# Get birthdays this week and this month
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
root.geometry("400x250")
root.config(bg="#FFFFFF")

# Title
title = tk.Label(root, text="Upcoming Birthdays", font=("Helvetica", 20, "bold"), fg="#333333", bg="#FFFFFF")
title.pack(pady=20)

# Birthday this week
if this_week:
    this_week_frame = tk.Frame(root, bg="#FFFFFF")
    this_week_frame.pack(fill="both", expand=True)
    this_week_label = tk.Label(this_week_frame, text="This Week", font=("Helvetica", 14, "bold"), fg="#333333", bg="#FFFFFF")
    this_week_label.pack(pady=5)
    for name in this_week:
        message = f"{name} has a birthday this week!"
        tk.Label(this_week_frame, text=message, font=("Helvetica", 12), fg="#333333", bg="#FFFFFF").pack()

# Birthday this month
if this_month:
    this_month_frame = tk.Frame(root, bg="#FFFFFF")
    this_month_frame.pack(fill="both", expand=True)
    this_month_label = tk.Label(this_month_frame, text="This Month", font=("Helvetica", 14, "bold"), fg="#333333", bg="#FFFFFF")
    this_month_label.pack(pady=5)
    for name in this_month:
        message = f"{name} has a birthday this month!"
        tk.Label(this_month_frame, text=message, font=("Helvetica", 12), fg="#333333", bg="#FFFFFF").pack()

root.mainloop()
