from openpyxl import Workbook

# Create a new workbook
wb = Workbook()

# Select the active sheet
sheet = wb.active

# Rename sheet
sheet.title = "Accounts"

# Create headings
sheet.append([
    "Account No",
    "Name",
    "Contact",
    "PIN",
    "Balance"
])

# Save workbook
wb.save("PROJECT.xlsx")

print("Database Created Successfully")
