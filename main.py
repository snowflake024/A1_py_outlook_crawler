import datetime
import win32com.client as win32
import re

# today_date = str(datetime.date.today())
outlook = win32.Dispatch('outlook.application').GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
message = inbox.Folders.Item("KPI")
kpi_messages = message.Items

if len(kpi_messages) == 0:
    print("There aren't any messages in this folder")
    exit()

emails = []
i = 0
c = 0

for message in kpi_messages:
    # get some information about each message in a tuple
    i += 1
    this_message = (
        message.Subject,
        message.SenderEmailAddress,
        message.To,
        message.Unread,
        message.Senton.date(),
        message.body,
        message.Attachments
    )
    # add this tuple of info to a list holding the messages
    emails.append(this_message)
    split_from_address = message.SenderEmailAddress.partition("-")[2]
    
    # show the results
    print(message.Subject, message.Senton.date(), split_from_address)


