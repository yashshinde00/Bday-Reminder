import pandas as pd
from datetime import date
from plyer import notification


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


        print(f"Today is {name}'s birthday! ")
        if __name__ == "__main__":
            notification.notify(
                    title = "Birthday remainder",
                    message = f"Today is {name}'s birthday! " ,
                    timeout = 25
                )
            


            