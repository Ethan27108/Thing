from openpyxl import load_workbook

# Load your workbook and target sheet
wb = load_workbook("your_file.xlsx")
ws = wb["parts"]

# Normalize your part number for matching
part_number = "ABC123"

# Find the header row and column indices for AS400, V, and X
headers = {}
for col in ws.iter_cols(min_row=1, max_row=1):
    header = col[0].value
    if header is not None:
        headers[header.strip()] = col[0].column  # openpyxl column index (int)

# Check needed columns exist
for col_name in ["AS400", "Location", "Status"]:  # Replace Location and Status with your actual column headers for V and X
    if col_name not in headers:
        raise ValueError(f"Column '{col_name}' not found!")

as400_col = headers["AS400"]
location_col = headers["Location"]  # Replace if needed
status_col = headers["Status"]      # Replace if needed

# Iterate rows to find matching part number (in AS400 column)
found = False
for row in ws.iter_rows(min_row=2):  # skip header row
    cell_value = str(row[as400_col - 1].value).strip()  # openpyxl columns are 1-based, list is 0-based
    if cell_value == part_number:
        # Update columns V and X (Location and Status here)
        row[location_col - 1].value = "New Location Value"
        row[status_col - 1].value = "New Status Value"
        found = True
        break

if not found:
    raise ValueError(f"Part number '{part_number}' not found.")

# Save workbook back, preserving filters and formatting
wb.save("your_file.xlsx")
print("✅ Update complete and formatting preserved.")
