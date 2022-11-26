from dataclasses import dataclass

# Ref https://xlsxwriter.readthedocs.io/index.html
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

# Pandas is great, but I think a big powerful library we don't need yet.
# Ref: https://medium.com/@AIWatson/how-to-read-csv-files-in-python-without-pandas-b693fc7ea3b7 and then
# https://docs.python.org/3/library/csv.html for DictReader
import csv
import sys

# Initial version of this script takes a single file as input
filename = sys.argv[1]

# Input is expected to be CSV in following form (as of this writing):

# Date,Time,Time Zone,Gross Sales,Discounts,Service Charges,Net Sales,Gift Card Sales,Tax,Tip,Partial Refunds,Total Collected,Source,Card,Card Entry Methods,Cash,Square Gift Card,Other Tender,Other Tender Type,Other Tender Note,Fees,Net Total,Transaction ID,Payment ID,Card Brand,PAN Suffix,Device Name,Staff Name,Staff ID,Details,Description,Event Type,Deposit ID,Location,Dining Option,Fee Percentage Rate,Fee Fixed Rate,Refund Reason,Discount Name,Transaction Status,Cash App
# 2022-11-10,21:50:48,Pacific Time (US & Canada),$22.50,$0.00,$0.00,$22.50,$0.00,$0.00,$0.00,$0.00,$22.50,eCommerce Integrations,$22.50,Keyed,$0.00,$0.00,$0.00,"","",-$0.95,$21.55,H0gwRVidtViqrKOeFk52Ogi0gZTZY,F8r3g2ZWmeSIbkg0tdOH7gCjiFAZY,Visa,7685,,,,https://squareup.com/dashboard/sales/transactions/H0gwRVidtViqrKOeFk52Ogi0gZTZY/by-unit/L08F63EFNAXA2,"Custom Amount - Reservation for Fabrication Access PM, Sat 11/12 6:00pm, full name: FOO BAR, phone: 2175551212, email: foo.bar@outlook.com",Payment,3ZMFS7QXV27W8YVCV8BY12E3434N,Pratt Fine Arts Center,"",2.9,$0.30,"","",Complete,$0.00


# Our output will hide all the other than the ones in the following include-list.
shown_column_names: tuple[str] = ("Date", "Card", "Fees", "Net Total", "Description")

summed_column_names: tuple[str] = "Net Total"

# A helper method to ensure that these columns are present
def validate(row: dict):
    for c in shown_column_names:
        if c not in row:
            raise f"Column {c} not found in columns from {filename}"


# GL "dataclass"
@dataclass
class GL:
    id: str
    name: str


# Hardcoded GL mapping
# I reordered it a bit (from what we got from the HTML source)
# Ideas for making this easier: For each description, sort them
# by descending length of longest common substring with the description.
# Then the correct user input will almost always be "1"
gls: tuple[GL] = (
    GL(id="411601512", name="Printmaking"),
    GL(id="411601514", name="Color Processor"),
    GL(id="411601521", name="Hot Glass"),
    GL(id="411601523", name="Flat Glass"),
    GL(id="411601524", name="Flameworking"),
    GL(id="411601531", name="Wood"),
    GL(id="411601534", name="Blacksmithing/Forgin"),
    GL(id="411601535", name="Fabrication"),
    GL(id="411601540", name="Jewelry"),
    GL(id="411601522", name="Coldworking"),
    GL(id="411601350100", name="2D"),
    GL(id="411601530", name="Sculpture"),
)


def getGL(description: str):
    print()
    print("GL choices:")
    i = 1
    for gl in gls:
        print(f"[{i}] {gl.name}")
        i = i + 1

    print()
    print("Description:")
    # Clean up the description, which itself looks like generated CSV
    dparts = description.split(",")
    print(dparts[0])
    print()
    while True:
        try:
            user_input = input("Choose GL by number: ")
            # Convert it into integer
            val = int(user_input)
            if val >= 1 and val <= len(gls):
                return gls[val - 1]
            else:
                print(f"Value must be between 1 and {len(gls)}, inclusive")
        except ValueError:
            print("Input must be an integer.  Try again")


# So long as we really want a workook per day (rather than, say, a worksheet per day in a single workbook),
# we keep track of just the list of all workbooks we've opened, plus a dictionary of worksheets per filename
# All the spreadsheet related stuff
@dataclass
class SSheet:
    workbook: xlsxwriter.Workbook
    worksheet: xlsxwriter.Workbook.worksheet_class
    row_count: int
    currency_format: int


ssheets: dict[str, SSheet] = {}


def appendWorksheetRow(
    ssheet: SSheet,
    row: dict,
    gl: str,
):
    i = 0
    worksheet = ssheet.worksheet
    row_count = ssheet.row_count
    for key, val in row.items():
        worksheet.write(row_count, i, val)
        if key == "Date":
            worksheet.write(row_count + 1, i, " ")
            worksheet.write(row_count + 2, i, " ")
            worksheet.write(row_count + 3, i, "Totals:")
        if key in summed_column_names:
            # Rewrite as a number so the formula calculation will work.
            worksheet.write_number(row_count, i, float(str.replace(val, '$', '')), ssheet.currency_format)
            worksheet.write(row_count + 1, i, " ")
            worksheet.write(row_count + 2, i, " ")
            formula = (
                f"SUM({xl_col_to_name(i)}2:{xl_col_to_name(i)}{row_count+1})"
            )
            print(formula)
            worksheet.write_formula(
                row_count + 3, i, formula, value="10"
            )
        i += 1
    worksheet.write(row_count, i, gl)
    ssheet.row_count += 1


def updateWorksheet(date: str, row: dict, gl: GL):
    if date not in ssheets:
        workbook = xlsxwriter.Workbook(f"{date}.xlsx")
        worksheet = workbook.add_worksheet()
        # Create columns
        i = 0
        for key in row:
            worksheet.write(0, i, key)
            worksheet.set_column(
                first_col=i,
                last_col=i,
                width=50 if key == "Description" else 10,
                options={"hidden": key not in shown_column_names},
            )
            i += 1
        worksheet.write(0, i, "GL")
        # A thing we need to do
        # Add to global collection
        ssheets[date] = SSheet(
            workbook=workbook,
            worksheet=worksheet,
            row_count=1,
            currency_format=workbook.add_format({"num_format": "$#,##0.00"}),
        )
        # End creation of new workbook & worksheet for the date in question
    appendWorksheetRow(ssheets[date], row, gl)


with open(filename, "r") as csvfile:
    reader_variable = csv.DictReader(csvfile, delimiter=",")
    for row in reader_variable:
        validate(row)
        date = row["Date"]
        gl:GL = getGL(row["Description"])
        worksheet: xlsxwriter.Workbook.worksheet_class = updateWorksheet(
            date, row, gl.id
        )
        print(f"Thanks!  Added row number {ssheets[date].row_count} to {date}.xls")
        print(f"------------------------------------------------------")

for ssheet in ssheets.values():
    ssheet.workbook.close()
