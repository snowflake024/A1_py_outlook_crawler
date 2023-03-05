from fetch_email import *
from import_data import *

# invocation of first part
print("")
print(f"Today's date is {today_date}\n")
print("++++++++++++++++++++++++++++++++")
print_all_emails()
print("\n++++++++++++++++++++++++++++++++")
input_date = str(input("Please choose an email's date to be processed: "))
#input_date = "2022-01-20"
fetch_email_attachment(input_date)


# topaz data
data_to_import = get_service_data(current_week_v1())
# current week v2
c_week = current_week_v2(data_to_import, kpi_ws_months())

# invocation of second part
import_service_data(data_to_import, c_week, kpi_ws_months)
