import win32com.client
import os

# Initialize the Outlook application
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Choose the mailbox
inbox = outlook.Folders.Item("support@portofwaterford.com").Folders.Item("Inbox")

# Folder to save logs
save_folder = os.path.expanduser("~/Desktop/AntivirusLogs")
os.makedirs(save_folder, exist_ok=True)

# Iterate through the emails (check 2CB8ED7C7280 this is where subject is written)
messages = inbox.Items
for message in messages:
    if message.Class == 43: # Check if the item is a MailItem
        if "2CB8ED7C7280" in message.Subject: # Check if the subject matches
            attachments = message.Attachments
            for attachment in attachments:
                attachment.SaveAsFile(os.path.join(save_folder, attachment.FileName))
                print(f"Saved attachment: {attachment.FileName}")

print("Done.")
