import win32com.client
import re

def minimize_body(body):
    """
    Minimize email body by removing signatures, replies, and excess whitespace.
    """
    # Remove signatures based on common patterns
    signature_patterns = [
        r'(--\s)',          # Signature starting with --
        r'(Best regards,)', # Common sign-off phrase
        r'(Thanks,)',       # Another sign-off
        r'(Sincerely,)',    # Another sign-off
        r'(Sent from my)',  # Mobile signature
        r'\bE:\s.*',  # Lines starting with 'E:' for email addresses
        r'<mailto:.*?>',  # Email addresses in <mailto:...> format
        r'https?:\/\/\S+',  # URLs
        r'www\.\S+',  # Web links starting with www.
        r'Vladimira Popovića.*',  # Specific address (customizable)
        r'[\r\n]+_{2,}',  # Lines with long underscores or separators
    ]
    for pattern in signature_patterns:
        match = re.search(pattern, body, re.IGNORECASE)
        if match:
            body = body[:match.start()]
            break

    # Remove quoted replies (e.g., "From:", "Sent:", etc.)
    reply_patterns = [
        r'From:\s',  # Quoted reply indicator
        r'Sent:\s',  # Outlook reply format
        r'Original Message',  # Forwarded/replied content
    ]
    for pattern in reply_patterns:
        match = re.search(pattern, body, re.IGNORECASE)
        if match:
            body = body[:match.start()]
            break

    # Remove excessive blank lines
    body = "\n".join([line.strip() for line in body.splitlines() if line.strip()])
    return body.strip()

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
                
                # Minimize the email body
                body = minimize_body(message.Body)
                file.write(f"Body:\n{body}\n")
                file.write("\n" + "="*100 + "\n\n")  # Separator between emails

    print("Email thread saved to email_thread.txt.")
