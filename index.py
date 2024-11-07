import win32com.client

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Access the Inbox
inbox = outlook.GetDefaultFolder(6)  # 6 refers to the Inbox folder

# Prompt the user to enter the conversation title they want to search for
conversation_title = input("Enter the conversation title to read: ")

# Define the signature start keyword to identify where the signature begins
signature_start = "Dušica Drača"

# Get all items (emails) in the Inbox
messages = inbox.Items

# Filter emails by conversation title
for message in messages:
    if conversation_title.lower() in message.Subject.lower():
        print(f"Subject: {message.Subject}")
        print(f"From: {message.SenderEmailAddress}")
        print(f"Received: {message.ReceivedTime}")
        
        # Truncate the body from the start of the signature if found
        body = message.Body
        if signature_start in body:
            body = body.split(signature_start)[0]
        
        print(f"Body:\n{body}\n")
