import datetime
import win32com.client as win32

today_date = str(datetime.date.today())
outlook = win32.Dispatch('outlook.application').GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
message = inbox.Folders.Item("KPI")
kpi_messages = message.Items

if len(kpi_messages) == 0:
    print("There aren't any messages in this folder")
    exit()

emails = []

print(f"Today date is {today_date}")
date_input = input("Please insect a date (follow above format): ")

this_message = ()

for message in kpi_messages:
    # get some information about each message in a tuple
    this_message = (
        message.Subject,
        message.SenderEmailAddress,
        message.To,
        message.Unread,
        message.Senton.date(),
        message.body,
        message.Attachments
    )


    emails.append(this_message)
    split_from_address = message.SenderEmailAddress.partition("-")[2]

    # Print relevant email entry by date
    # Add option to see all the emails available
    if str(message.Senton.date()) == date_input:
        print("am I being checked?")
        print(message.Subject, message.Senton.date(), split_from_address)
        exit()

