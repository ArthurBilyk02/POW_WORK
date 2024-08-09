import win32com.client
import os

# Email credentials
target_sender = "support@portofwaterford.com"
target_subject = "This is a test 1"

# Initialize the Outlook application
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Access the Inbox folder of the default mailbox
inbox = outlook.Folders.Item("support@portofwaterford.com").Folders.Item("Inbox")


# Folder to save logs
save_folder = os.path.expanduser("~/Desktop/AntivirusLogs")
os.makedirs(save_folder, exist_ok=True)

# Iterate through the emails
messages = inbox.Items
for message in messages:
    if message.Class == 43: # Check if the item is a MailItem
        if (message.Subject and target_subject in message.Subject) and (message.SenderEmailAddress and target_sender in message.SenderEmailAddress):
            attachments = message.Attachments
            for attachment in attachments:
                attachment.SaveAsFile(os.path.join(save_folder, attachment.FileName))
                print(f"Saved attachment: {attachment.FileName}")

print("Done.")
