from dataclasses import dataclass

# Ref https://xlsxwriter.readthedocs.io/index.html
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name
import operator

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

summed_column_names: tuple[str] = ("Card", "Fees", "Net Total")

# A helper method to ensure that these columns are present
def validate(row: dict):
    for c in shown_column_names:
        if c not in row:
            raise Exception(f"Column {c} not found in columns from {filename}")


# GL "dataclass"
@dataclass
class GL:
    id: str
    name: str
    # Mutable score field used for sorting based on LCS w/ description
    score: int = 0


# Hardcoded GL mapping
# We use a mutable score field to sort according to
# lenght of longest common substring with any given description
gls: tuple[GL] = (
    GL(id="411601512", name="Printmaking"),
    GL(id="411601514", name="Color Processor"),
    GL(id="411601521", name="Hot Glass"),
    GL(id="411601523", name="Flat Glass"),
    GL(id="411601524", name="Flameworking"),
    GL(id="411601531", name="Woodworking"),
    GL(id="411601534", name="Blacksmithing/Forging"),
    GL(id="411601535", name="Fabrication"),
    GL(id="411601540", name="Jewelry and Metalsmithing"),
    GL(id="411601522", name="Coldworking"),
    GL(id="411601350100", name="2D"),
    GL(id="411601530", name="Sculpture"),
)


# https://www.geeksforgeeks.org/longest-common-substring-dp-29/
def LCSubStr(X, Y, m, n):

    # Create a table to store lengths of
    # longest common suffixes of substrings.
    # Note that LCSuff[i][j] contains the
    # length of longest common suffix of
    # X[0...i-1] and Y[0...j-1]. The first
    # row and first column entries have no
    # logical meaning, they are used only
    # for simplicity of the program.

    # LCSuff is the table with zero
    # value initially in each cell
    LCSuff = [[0 for k in range(n + 1)] for l in range(m + 1)]

    # To store the length of
    # longest common substring
    result = 0

    # Following steps to build
    # LCSuff[m+1][n+1] in bottom up fashion
    for i in range(m + 1):
        for j in range(n + 1):
            if i == 0 or j == 0:
                LCSuff[i][j] = 0
            elif X[i - 1] == Y[j - 1]:
                LCSuff[i][j] = LCSuff[i - 1][j - 1] + 1
                result = max(result, LCSuff[i][j])
            else:
                LCSuff[i][j] = 0
    # print(f"lcs of {X} and {Y} is {result}")
    return result


# Sorts the GLs based on lenght of longest substring w/ description
def sort_gls(description: str):
    for gl in gls:
        gl.score = LCSubStr(gl.name, description, len(gl.name), len(description))
    return [e for e in reversed(sorted(gls, key=operator.attrgetter("score")))]


# The descriptions have a bunch of boilerplate and additional info we don't care about
def clean_up_description(desc: str):
    return (
        desc.split(",")[0]
        .replace("Custom Amount - Reservation for ", "")
        .replace(" Studio Access PM", "")
        .replace(" Access PM", "")
    )


def getGL(description: str):
    # Clean up the description, which itself looks like generated CSV
    desc = clean_up_description(description)
    gls_sorted = sort_gls(desc)
    print()
    print(f"De: {desc}")
    i = 1
    for gl in gls_sorted:
        print(f"[{i}] {gl.name} ({gl.score})")
        i = i + 1

    while True:
        try:
            user_input = input("Choose GL by number [1]: ")
            user_input = 1 if len(user_input) == 0 else user_input
            # Convert it into integer
            val = int(user_input)
            if val >= 1 and val <= len(gls_sorted):
                return gls_sorted[val - 1]
            else:
                print(f"Value must be between 1 and {len(gls)}, inclusive")
        except ValueError:
            print("Input must be an integer.  Try again")


# So long as we really want a workook per day (rather than, say, a worksheet per day in a single workbook),
# we need to have a workbook object per worksheet.  We also need to keep track of the row_count, and
# just to be safe, each workbook/worksheet has its own currency_format object
@dataclass
class SSheet:
    workbook: xlsxwriter.Workbook
    worksheet: xlsxwriter.Workbook.worksheet_class
    row_count: int
    currency_format: any


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
            # worksheet.write(row_count, i, val)
            # Rewrite as a number so the formula calculation will work.
            curr_num = float(str.replace(str.replace(val, "$", ""), ",", ""))
            worksheet.write_number(row_count, i, curr_num, ssheet.currency_format)
            worksheet.write(row_count + 1, i, " ")
            worksheet.write(row_count + 2, i, " ")
            # Will end up as 'SUM(V2:V6)'
            formula = f"SUM({xl_col_to_name(i)}2:{xl_col_to_name(i)}{row_count+1})"
            worksheet.write_formula(row_count + 3, i, formula, ssheet.currency_format)
        i += 1
    worksheet.write_number(row_count, i, int(gl))
    ssheet.row_count += 1


def updateWorksheet(date: str, row: dict, gl: GL):
    if date not in ssheets:
        workbook = xlsxwriter.Workbook(f"{date}.xlsx")
        worksheet = workbook.add_worksheet()
        # Create column headers
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
        worksheet.set_column(first_col=i, last_col=i, width=10)

        # Add the stuff to our global collection of spreadsheets (ssheets)
        ssheets[date] = SSheet(
            workbook=workbook,
            worksheet=worksheet,
            row_count=1,
            currency_format=workbook.add_format({"num_format": "$#,##0.00"}),
        )
        # End creation of new workbook & worksheet for the date in question
    appendWorksheetRow(ssheets[date], row, gl)


# Main portion of the program
with open(filename, "r") as csvfile:
    reader = csv.DictReader(csvfile, delimiter=",")
    for row in reader:
        print(f"------------------------------------------------------")
        validate(row)
        date = row["Date"]
        gl: GL = getGL(row["Description"])
        print(f"We picked gl = {gl.id}, {gl.name}")
        updateWorksheet(date, row, gl.id)
        print(f"Thanks!  Added row number {ssheets[date].row_count} to {date}.xls")

print()
for name, ssheet in ssheets.items():
    print(f"Writing excel file {name}")
    ssheet.workbook.close()
