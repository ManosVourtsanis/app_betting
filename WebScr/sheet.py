import gspread
from oauth2client.service_account import ServiceAccountCredentials
from prettytable import PrettyTable
import pandas as pd

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

creds = ServiceAccountCredentials.from_json_keyfile_name("C:\\Users\\Basilis\\Desktop\Python\\New-betting_app\\secret_key.json", scopes=scopes)

file = gspread.authorize(creds)
workbook = file.open("Bets by Tolis")
sheet = workbook.sheet1

def get_valid_integer(prompt):
    while True:
        try:
            value = int(input(prompt))
            return value
        except ValueError:
            print("Please enter a valid number.")


def get_valid_float(prompt):
    while True:
        try:
            value = float(input(prompt).replace(',', '.'))
            return value
        except ValueError:
            print("Please enter a valid number.")


def find_first_empty_row(sheet, column='A'):
    """Find the first empty row in the specified column."""
    col_index = gspread.utils.a1_to_rowcol(column + '1')[1]  # Convert column letter to index
    col_values = sheet.col_values(col_index)
    for i, value in enumerate(col_values):
        if not value.strip():
            return i + 1
    return len(col_values) + 1


def view_data():
    try:
        data = sheet.get_all_values()
        if not data:
            print("No data found.")
            return
        
        # Create a PrettyTable object
        table = PrettyTable()
        # Set the table headers
        table.field_names = data[0]
        # Add rows to the table
        for row in data[1:]:
            table.add_row(row)

        # Print the table
        print(table)
    except Exception as e:
        print(f"An error occurred: {e}")


def insert_data():
    # Collect data for one row
    date = input("Enter Date (DD-MM-YYYY): ")
    
    num_matches = get_valid_integer("How many matches do you want to add (1-5)? ")
    while num_matches < 1 or num_matches > 5:
        num_matches = get_valid_integer("How many matches do you want to add (1-5)? ")
    
    matches = []
    odds = []

    for i in range(num_matches):
        match = input(f"Enter Match No{i+1}: ")
        odd = get_valid_float(f"Enter Odd_{i+1}: ")
        matches.append(match)
        odds.append(odd)

    # Append empty strings for matches and odds if fewer than 5 matches
    while len(matches) < 5:
        matches.append("")
    while len(odds) < 5:
        odds.append("")

    stake = get_valid_float("Enter Stake: ")

    # Default result is "Pending"
    result = "Pending"

    # Append row data to the data list
    data = [
        date,
        matches[0], odds[0],
        matches[1], odds[1],
        matches[2], odds[2],
        matches[3], odds[3],
        matches[4], odds[4],
        stake, result
    ]

    # Find the first empty row in the "Date" column (Column A)
    first_empty_row = find_first_empty_row(sheet, 'A')

    # Write the data to the sheet starting from the first empty row
    for j, value in enumerate(data, start=1):
        if j not in [14, 15, 16]:  # Avoid writing to columns with formulas
            sheet.update_cell(first_empty_row, j, value) 

    print("Data successfully inserted.")

"""
def delete_row():
    # Prompt user for the row number to clear
    row_to_delete = get_valid_integer("Enter the row number you want to delete (2 to {}): ".format(sheet.row_count))

    # Validate the row number
    if row_to_delete < 2 or row_to_delete > sheet.row_count:
        print(f"Invalid row number. Please enter a number between 2 and {sheet.row_count}.")
        return

    # Define columns to ignore (e.g., columns with formulas)
    ignore_columns = [14, 15, 16]  # Columns to be preserved

    # Prepare a list of update requests for batch update
    update_requests = []

    # Get the total number of columns in the sheet
    max_col = sheet.col_count

    # Clear the row (excluding columns to be preserved)
    for col in range(1, max_col + 1):
        if col not in ignore_columns:
            update_requests.append({
                'range': f'{gspread.utils.rowcol_to_a1(row_to_delete, col)}',
                'values': [['']]
            })

    # Shift up the rows below
    for row in range(row_to_delete + 1, sheet.row_count + 1):
        for col in range(1, max_col + 1):
            if col not in ignore_columns:
                value_above = sheet.cell(row, col).value
                update_requests.append({
                    'range': f'{gspread.utils.rowcol_to_a1(row - 1, col)}',
                    'values': [[value_above]]
                })
                # Clear the current cell
                update_requests.append({
                    'range': f'{gspread.utils.rowcol_to_a1(row, col)}',
                    'values': [['']]
                })

    # Execute batch update
    sheet.batch_update(update_requests)

    print(f"Row {row_to_delete} successfully deleted and remaining rows shifted up.")
"""
#find_first_empty_row(sheet, 'A'):

def delete_row():
    # Get the total number of rows
    num_rows = len(sheet.get_all_values())  # Total number of rows in the sheet
    
    # Prompt user for the row number to clear
    row_to_clear = get_valid_integer(f"Enter the row number to clear (1-{num_rows}): ")
    
    # Validate the row number
    if 1 <= row_to_clear <= num_rows:
        # Clear the first 13 cells in the specified row
        for col in range(1, 14):  # Columns 1 through 13
            sheet.update_cell(row_to_clear, col, '')  # Set the cell to empty
        
        # Fetch the data from rows below
        for row in range(row_to_clear, num_rows):
            # Get the next row values
            row_values = sheet.row_values(row + 1)
            
            # Prepare new row values by limiting to first 13 columns and preserving columns 14-16
            if len(row_values) > 13:
                new_row_values = row_values[:13]
            else:
                new_row_values = row_values + [''] * (13 - len(row_values))

            # Insert the new row values into the current row position
            # Only inserting the first 13 columns
            sheet.insert_row(new_row_values, row + 1)
            
            # Delete the last row to keep the number of rows consistent
            sheet.delete_rows(num_rows + 1)
        
        print(f"Cleared the first 13 cells in row {row_to_clear} and shifted up the rows below.")
    else:
        print(f"Invalid row number. Please enter a number between 1 and {num_rows}.")

def menu():
    while True:
        print("\nMenu:")
        print("1. View Data")
        print("2. Insert Data")
        print("3. Delete Row")
        print("4. Exit")

        choice = get_valid_integer("Enter your choice (1-4): ")

        if choice == 1:
            view_data()
        elif choice == 2:
            insert_data()
        elif choice == 3:
            delete_row()
        elif choice == 4:
            print("Exiting...")
            break
        else:
            print("Invalid choice. Please enter a number between 1 and 4.")

if __name__ == "__main__":
    menu()
