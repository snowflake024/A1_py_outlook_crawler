import datetime
import os
import win32com.client as win32

#folder = str(input("Choose a folder: "))
folder = "KPI"
path = r"C:\Users\a1bg482511\Desktop\lib\Repo\PYTHON\email_parser\attachments"
today_date = str(datetime.date.today())

# Variables related to win32com.client
outlook = win32.Dispatch('outlook.application').GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
message_object = inbox.Folders.Item(folder)
items_messages = message_object.Items


def print_all_emails():
    """Print all emails in a given outlook folder"""

    if len(items_messages) == 0:
        print("There aren't any messages in this folder")
        exit()

    all_emails = []
    for message in items_messages:
        # get some information about each message in a tuple
        this_message = (
            message.Subject,
            message.SenderEmailAddress.partition("-")[2],
            message.To,
            message.Unread,
            message.Senton.date(),
            message.body,
            message.Attachments
        )

        all_emails.append(this_message)

    print("List of all emails in the chosen folder: \n")

    for i in all_emails:
        print(f"Subject: {i[0]}, Sender: {i[1]}, Sent on: {i[4]}")

    print("\n++++++++++++++++++++++++++++++++")
    print(f"List of available dates:\n")
    for message in items_messages:
        print(message.Senton.date())


def fetch_email_attachment(selected_date):
    """fetches a single email attachment to be imported in the KPI Excel spreadsheet"""

    for message in items_messages:

        if selected_date == str(message.Senton.date()):
            print(f"\nsaving attachments for email with Subject: {message.Subject}\n")

            for attachment in message.Attachments:
                print(f"{attachment.FileName}")
                if os.path.splitext(attachment.FileName)[1] == ".xlsx":
                    attachment.SaveAsFile(os.path.join(path, str(attachment.FileName)))

########################################################################################

print("")
print(f"Today's date is {today_date}\n")
print("++++++++++++++++++++++++++++++++")
print_all_emails()
print("\n++++++++++++++++++++++++++++++++")
#input_date = str(input("Please choose an email's date to be processed: "))
input_date = "2022-01-20"
fetch_email_attachment(input_date)

