from fetch_email import *

print("")
print(f"Today's date is {today_date}\n")
print("++++++++++++++++++++++++++++++++")
print_all_emails()
print("\n++++++++++++++++++++++++++++++++")
#input_date = str(input("Please choose an email's date to be processed: "))
input_date = "2022-01-20"
fetch_email_attachment(input_date)
