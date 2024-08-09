import win32com.client
import os

# Email credentials
target_sender = "support@portofwaterford.com"
target_subject = "This is a test"

# Initialize the Outlook application
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Access the Inbox folder of the specified mailbox
inbox = outlook.Folders.Item("support@portofwaterford.com").Folders.Item("Inbox")

# Folder to save logs
save_folder = os.path.expanduser("~/Desktop/AntivirusLogs")
os.makedirs(save_folder, exist_ok=True)

# Iterate through the emails
messages = inbox.Items
for message in messages:
    if message.Class == 43: # Check if the item is a MailItem
        sender_email = target_sender
        subject = target_subject

        print(f"Checking email from {sender_email} with subject '{subject}'") # Debug statement

        if subject and target_subject in subject and sender_email and target_sender in sender_email:
            print(f"Matched email from {sender_email} with subject '{subject}'") # Debug statement
            attachments = message.Attachments
            for attachment in attachments:
                attachment.SaveAsFile(os.path.join(save_folder, attachment.FileName))
                print(f"Saved attachment: {attachment.FileName}")

print("Done.")


Sent from Outlook for Android
