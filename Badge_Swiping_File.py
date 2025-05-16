import pandas as pd
import os
import sys
import subprocess

try:
    # Load the Excel file with a header row
    input_file = r"C:\Users\Rputhoff\Documents\Badge_Swiping\Cardholders with Active Cards ().xls"
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"Input file not found: {input_file}")
    
    # Check if the file is a valid Excel file
    if not input_file.endswith(('.xls', '.xlsx')):
        raise ValueError(f"Invalid file format: {input_file}. Expected an Excel file with .xls or .xlsx extension.")
    
    # Use appropriate engine based on file extension
    engine = 'xlrd' if input_file.endswith('.xls') else 'openpyxl'
    try:
        df = pd.read_excel(input_file, engine=engine)
    except ImportError as e:
        if engine == 'xlrd':
            print("Error: Missing optional dependency 'xlrd'. Attempting to install it...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", "xlrd>=2.0.1"])
            print("xlrd installed successfully. Please re-run the script.")
            sys.exit(1)
        else:
            raise

    # Delete specified columns based on their header names
    columns_to_delete = ["Last Name", "First Name", "Company Name", "Access Group", "Imprint", "Card Status", "CardholderType"]
    df = df.drop(columns=columns_to_delete)

    # Rename columns
    df = df.rename(columns={"Middle Name": "EmpID", "Cardnumber": "Badge_num"})

    # Remove rows where "Emp ID" is NULL
    df = df[df["EmpID"].notnull()]

    # Load the valid "Emp ID" values from the second file
    valid_file = r"C:\Users\Rputhoff\Documents\Badge_Swiping\Midmark Production Badge Swiping Group.xlsx"
    valid_df = pd.read_excel(valid_file, usecols=[0], engine='openpyxl')  # Only load Column A
    valid_emp_ids = valid_df.iloc[:, 0].dropna().astype(str).tolist()

    # Keep rows where "Emp ID" matches the valid list
    df = df[df["EmpID"].astype(str).isin(valid_emp_ids)]

    # Add leading zeros to "Badge Num" values that are not 10 digits
    df["Badge_num"] = df["Badge_num"].astype(str).apply(lambda x: x.zfill(10) if len(x) < 10 else x)

    # Save the result as a comma-delimited text file
    output_file = r"C:\Users\Rputhoff\Documents\Badge_Swiping\CS_badge_import.txt"
    df.to_csv(output_file, index=False, encoding='utf-8', sep=',')

except FileNotFoundError as fnf_error:
    print(f"Error: {fnf_error}")
except ValueError as val_error:
    print(f"Error: {val_error}")
except ImportError as imp_error:
    print(f"Error: {imp_error}")
except Exception as e:
    print(f"Unexpected error: {e}")
