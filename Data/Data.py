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

    # Read specific rows (e.g., rows 5 to 10) from the decrypted file
    df = pd.read_excel(
        decrypted,
        usecols="C",      
        nrows=9,
        engine='openpyxl'
    )
    df = df["Min"]

    df_1 = pd.read_excel(
        decrypted,
        usecols="D",      
        nrows=8,
        engine='openpyxl'
    )
    df_1 = df_1["Max"]

    # Combine Min and Max values
    min_max = []
    for i in range(0, 7):
        if df[i] == '-':
            min_max.append("NULL")
        else:
            min_max.append(str(df[i]))
        if df_1[i] == '-':
            min_max.append("NULL")
        else:
            min_max.append(str(df_1[i]))
    min_max.append(str(df[7]))
    # Print the result
    print(",".join(min_max))

    # Exit the program
    sys.exit(0)

if __name__ == "__main__":
    main()
