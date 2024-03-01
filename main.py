import win32com.client

def process_folder(main_folder_name, sub_folder_name, sheet_name):
    # Access the specified folder and subfolder in Outlook
    inbox = outlook.Folders[main_folder_name].Folders[sub_folder_name]
    # Access the specified worksheet in the Excel workbook
    ws = wb.Worksheets[sheet_name]
    # Determine the last row in the worksheet that has data
    last_row = ws.UsedRange.Rows.Count

    # Loop through each message in the inbox
    messages = inbox.Items
    for message in messages:
        categories = message.Categories
        # If the message has categories, process further
        if categories:
            subject = message.Subject
            received_time = message.ReceivedTime 
            sender_name = message.SenderName  
            # Write message details to the next available row in the worksheet
            ws.Cells(last_row + 1, 1).Value = subject
            ws.Cells(last_row + 1, 2).Value = categories
            ws.Cells(last_row + 1, 3).Value = received_time.Format("%Y-%m-%d %H:%M:%S")
            ws.Cells(last_row + 1, 4).Value = sender_name  
            # Increment the row counter
            last_row += 1

# Create COM object instances for Outlook and Excel
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
excel = win32com.client.Dispatch("Excel.Application")

# Replace the placeholder with the path to your Excel file
excel_file_path = r'<Path to your Excel file>'
wb = excel.Workbooks.Open(excel_file_path)

# Replace placeholder values with your actual folder names and sheet names
# Example placeholders are provided. Replace them with actual values before running the script.
process_folder('<Main folder name>', '<Subfolder name>', '<Sheet name>')
process_folder('<Another main folder name>', '<Another subfolder name>', '<Another sheet name>')

# Save the workbook and quit Excel
wb.Save()
excel.Quit()

# Clean up the COM objects
del outlook, wb, excel