import msoffcrypto
import pandas as pd
import io
import sys
import os

def main():
    # Specify the path to your Excel file
    file_path = os.path.join(os.getcwd(), 'Limits_data.xlsx')

    # Password to open the Excel file
    password = sys.argv[1]

    # Decrypt the file using msoffcrypto
    decrypted = io.BytesIO()
    with open(file_path, 'rb') as file:
        office_file = msoffcrypto.OfficeFile(file)
        office_file.load_key(password=password)
        office_file.decrypt(decrypted)

    # Read the Excel file from the decrypted data without headers
    df = pd.read_excel(decrypted, usecols="C:D", nrows=23, engine='openpyxl', header=None)
    
    # Assign column names since we have no headers
    df.columns = ["Min", "Max"]

    # Remove the first row
    df = df.iloc[1:].reset_index(drop=True)

    # Replace '-' with 'NULL' and NaNs with 'NULL'
    df = df.replace('-', 'NULL').fillna('NULL')

    # Combine the "Min" and "Max" columns into a single list
    min_max = []
    
    # Ensure to loop only over the actual length of the DataFrame
    for i in range(len(df)):
        min_val = df.at[i, "Min"]
        max_val = df.at[i, "Max"]
        
        min_max.append(str(min_val))
        min_max.append(str(max_val))

    # Print the result
    print(",".join(min_max))

    # Exit the program
    sys.exit(0)

if __name__ == "__main__":
    main()
