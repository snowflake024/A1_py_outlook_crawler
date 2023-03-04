from os import listdir
from list_of_services import list_of_services
import openpyxl as xl

# Avoid providing a certain name of the topaz report for simplicity
list_of_files = listdir(r"C:\Users\a1bg482511\Desktop\lib\Repo\PYTHON\email_parser\attachments")
filename = list_of_files[0]

# Paths to various involved files
path_to_KPI = r"C:\Users\a1bg482511\Desktop\lib\Repo\PYTHON\email_parser\KPI_test\KPI_REPORT-v2.0.xlsx"
path_to_Topaz = rf"C:\Users\a1bg482511\Desktop\lib\Repo\PYTHON\email_parser\attachments\{filename}"
path_to_services = r"C:\Users\a1bg482511\Desktop\lib\Repo\PYTHON\email_parser\list_of_services.txt"

# Removed the password protection of the KPI workbook because of the added complexity to read it
kpi_wb = xl.load_workbook(path_to_KPI)
topaz_wb = xl.load_workbook(path_to_Topaz, read_only=True)

kpi_sheets = kpi_wb.sheetnames
topaz_sheet = topaz_wb["Verfügbarkeit pro Woche(in %)"]


def get_current_week():
    """Check if the next cell in the KW* section is empty and return the previous one as current week"""
    # 5 + 52 column wise according to the topaz workbook
    week = 5
    while week < 57:
        if str(topaz_sheet.cell(2, week).value) == "None":
            week -= 1
            # it's better to return numerical value of column than cell contents
            #return topaz_sheet.cell(1, week).value
            return week
        else:
            week += 1

def get_service_data(kw):
    """Fetches the data for the current week"""

    # Create a dict with current week as first couple
    kpi_values = {'KW': topaz_sheet.cell(1, kw).value}

    for service in list_of_services:
        i = 0
        service_column = kw

        # Go through all rows in the first column to compare entries with service list
        for row in topaz_sheet.iter_rows():
            i += 1
            rows = f"A{i}"

            if service == topaz_sheet[rows].value:
                # print(topaz_sheet[rows].value)
                true_row = topaz_sheet[rows].row
                # Cell Cordinates in integers
                # print(f"{topaz_sheet[rows].row} {topaz_sheet[rows].column}")
                # Append all the services with their KPI % to the dict
                kpi_values[service] = topaz_sheet.cell(true_row, service_column).value
                break
    return kpi_values

print(get_service_data(get_current_week()))

