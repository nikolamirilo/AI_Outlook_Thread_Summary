import os
from groq import Groq


def summarizeEmailThreadData():
    client = Groq(api_key="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")

    file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "email_thread.txt")

    # Read the email thread content from the file
    with open(file_path, "r", encoding="utf-8") as file:
        input_data = file.read()

    # Define the prompt
    prompt = {
        "role": "you are IT project manager with more than 30 years of experience specialized in insurance industry",
        "task": "read email thread and summarize it",
        "format": "return in md format. Title of document (#) is subject, then mention who particpants are and then summarize conversation."
    }

    # Fix the f-string and reference the prompt values properly
    chat_completion = client.chat.completions.create(messages=[
        {
            "role": "system",
            "content": f"Role: {prompt['role']}"
        },
        {
            "role": "user",
            "content": f"Task: {prompt['task']} Here is the email thread: \n{input_data} | Format: {prompt['format']}"
        }
    ],
    model="llama3-8b-8192",
    temperature=0.1
    )

    # Extract the response content
    response = chat_completion.choices[0].message.content
    print(response)

    # Save the response to a file with utf-8 encoding
    save_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "response.md")

    try:
        with open(save_path, "w", encoding="utf-8") as file:  # Explicitly specify utf-8 encoding
            file.write(response)
        print(f"Successfully saved data to {save_path}")
    except Exception as e:
        print("Failed to save response:")
        print(e)
