import msoffcrypto
import pandas as pd
import io
import sys
import os
import win32com.client as win32

def main():
    # Specify the path to your Excel file
    file_path = os.path.join(os.getcwd(), 'Limits_data.xlsx')
    # Password to open the Excel file
    password = sys.argv[2].strip()

    # Decrypt the file
    decrypted = io.BytesIO()
    with open(file_path, 'rb') as file:
        office_file = msoffcrypto.OfficeFile(file)
        office_file.load_key(password=password)
        office_file.decrypt(decrypted)

    # Load the decrypted content into a DataFrame
    df = pd.read_excel(decrypted, engine='openpyxl')

    # Process the string and populate Min_L and Max_L lists
    string = sys.argv[1].split(',')
    print(string)
    weld=string[0:9]
    for i in range(0,9):
        print(string)
        string.pop(0)
        print(string)
    Min_L = []
    Max_L = []
    for i in weld:
        Max_L.append(i)
        Min_L.append("nan")
    for i, val in enumerate(string):
        if i % 2 == 0:
            Min_L.append(val)
        else:
            Max_L.append(val)

    # Assign the lists to the DataFrame columns
    Min_L.append(Max_L[-1])
    Max_L.pop(-1)
    Max_L.append("")
    Max_L.append("")
    print(Min_L,Max_L)
    df["Min"] = Min_L
    df["Max"] = Max_L

    # Replace NaN values with empty strings
    df.fillna('', inplace=True)

    # Save the modified DataFrame to a temporary file without a password
    temp_output_path = os.path.join(os.getcwd(), 'Temp_Limits_data.xlsx')
    df.to_excel(temp_output_path, index=False, engine='openpyxl', sheet_name='Sheet1')

    # Apply the password protection using win32com
    excel = win32.Dispatch('Excel.Application')
    excel.DisplayAlerts = False  # Suppress prompts like the "Replace" dialog

    workbook = excel.Workbooks.Open(temp_output_path)
    workbook.Password = password
    workbook.SaveAs(file_path, Password=password)
    workbook.Close(SaveChanges=True)
    
    excel.DisplayAlerts = True  # Re-enable alerts
    excel.Quit()

    # Remove the temporary file
    if os.path.exists(temp_output_path):
        os.remove(temp_output_path)

    print(f"Modified file saved to {file_path} with password protection.")

if __name__ == "__main__":
    main()
