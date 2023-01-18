import datetime
import win32com.client as win32

folder = str(input("Choose a folder: "))
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

    print("List of all emails in the chosen folder: ")

    for i in all_emails:
        print(f"Subject: {i[0]}, Sender: {i[1]}, Sent on: {i[4]}")


print("")
print(f"Today's date is {today_date}")
print("")
print_all_emails()
# choice = input("Please choose an email to be imported in the KPI Excel file")
