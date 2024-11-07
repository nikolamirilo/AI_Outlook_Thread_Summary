from extract_data import extractEmailThread
from summarize_data import summarizeEmailThreadData


def main():
    conversation_title = input("Enter the conversation title to read: ")
    extractEmailThread(conversation_title)  
    summarizeEmailThreadData() 
main()