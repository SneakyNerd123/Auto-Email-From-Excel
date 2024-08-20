import pandas as pd
import win32com.client as win32
import time

# Load the Excel file
file_path = 'REPLACEWITHYOURFILENAME.xlsx'  # Replace with the actual file path
df = pd.read_excel(file_path)

# Initialize Outlook application
olApp = win32.Dispatch('Outlook.Application')
olNS = olApp.GetNamespace('MAPI')

# Your email address for CC
cc_email = 'YourEmail@YourEmail'

# Loop through each row in the DataFrame
for index, row in df.iterrows():
    if row['Roommate Agreements'] == 'No':  # Adjust 'Roommate Agreements' to the actual column header
        first_name = row['First Name']  # Adjust 'First Name' to the actual column header
        last_name = row['Last Name']    # Adjust 'Last Name' to the actual column header
        email = row['Email']            # Adjust 'Email' to the actual column header
        
        # Create the email
        mailItem = olApp.CreateItem(0)  # 0: olMailItem
        mailItem.Subject = "Roommate Agreement Sign-Up"
        mailItem.BodyFormat = 1  # 1: Plain Text

        mailItem.Body = (
            f"Hello {first_name} {last_name},\n\n"
            "I noticed that you haven't signed up for a roommate agreement. "
            "This is a friendly reminder to sign up for it as we have limited spots. "
            "Please do so as soon as possible.\n\n"
            "The sign-up link can be found here: "
            "Thank you!\n\n"
            "Best,\nYour RA"
        )
        
        mailItem.To = email
        mailItem.CC = cc_email
        mailItem.SendUsingAccount = olNS.Accounts.Item('YourEmail@YourEmail')

        # Send the email
        try:
            mailItem.Send()
            print(f"Email sent to {email} with a CC to {cc_email}.")
        except Exception as e:
            print(f"Failed to send email to {email}. Error: {e}")
        
        # Add a small delay between sending emails
        time.sleep(2)

print("All emails have been processed.")
