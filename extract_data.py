import win32com.client

# Connect to Outlook
def extractEmailThread(conversation_title: str):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Access the Inbox
    inbox = outlook.GetDefaultFolder(6)  # 6 refers to the Inbox folder

    # Get all items (emails) in the Inbox
    messages = inbox.Items

    # Open a text file to save the email thread
    with open("email_thread.txt", "w", encoding="utf-8") as file:
        # Filter emails by conversation title
        for message in messages:
            if conversation_title.lower() in message.Subject.lower():
                # Write email details to the file
                file.write(f"Subject: {message.Subject}\n")
                file.write(f"From: {message.SenderEmailAddress}\n")
                file.write(f"Received: {message.ReceivedTime}\n")
                
                # Truncate the body from the start of the signature if found
                body = message.Body
                # Write the email body to the file
                file.write(f"Body:\n{body}\n")
                file.write("\n" + "="*250 + "\n\n")  # Separator between emails

    print("Email thread saved to email_thread.txt.")
