# The account payables in Xero shows all transaction history without ability to show only outstanding transactions.
# This script reads in the exported Account Payable in Excel format.
# Find pairs of matching debit and credit entries by value from the "Credit (Source)" and "Debit (Source)" columns enabling support for multi-currency accounts.
# It then generates 2 output files in the original directory of the input file:
# 1. Outstanding.xlsx : contains only outstanding payable items with the matches pairs removed.
# 2. removed.xlsx     : contains all the removed paired items for cross-checking, sum of each credit & debit columns should match.

# You can generate an *.exe file with pyinstaller in windows, *.exe will be generated in a "dist" directory 
# %> pyinstaller --onefile <this filename>
#
# You can then drag-and-drop your excel onto the *.exe file for processing.

#Copyright 2023 Jonathan Cheah
#Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files 
#the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, 
#publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do 
#so, subject to the following conditions:
#   Attribution to the original code owner
#
#The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
#THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES 
#OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE 
#LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR 
#IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.



import pandas as pd
import sys
import os

def process_excel(excel_path):
    # Read the Excel file
    df = pd.read_excel(excel_path, engine='openpyxl')

    # Drop rows 1-4 and 6-8 and reset the index
    df = df.drop(index=range(0, 3)).drop(index=range(5, 7)).reset_index(drop=True)

    # Reassign header based on the first row and drop the first row:
    df.columns = df.iloc[0]
    df = df.drop(0).reset_index(drop=True)

    # Print column names to check:
    #print(df.columns)
    
    #return 
    
    # Filter out zeros from 'Debit (Source)' and 'Credit (Source)' columns
    debits = df[df['Debit (Source)'] != 0]['Debit (Source)']
    credits = df[df['Credit (Source)'] != 0]['Credit (Source)']

    # Find pairs where 'Debit (Source)' in one row matches 'Credit (Source)' in another row (ignoring zeros)
    matched_pairs = set(debits).intersection(set(credits))

    # Filter out the rows that are part of these pairs to a separate DataFrame
    removed_rows = df[(df['Debit (Source)'].isin(matched_pairs)) | (df['Credit (Source)'].isin(matched_pairs))]
    
 
    # Filter out those paired rows from the original DataFrame to keep the rows we want to retain
    processed_rows = df[~df.index.isin(removed_rows.index)]

    # Save the paired rows to an Excel file named "removed.xlsx" in the same directory as the input file
    removed_output_path = os.path.join(os.path.dirname(excel_path), "removed.xlsx")
    removed_rows.to_excel(removed_output_path, index=False, engine='openpyxl')

    # Save the unpaired rows to a new Excel file named "processed.xlsx" in the same directory as the input file
    processed_output_path = os.path.join(os.path.dirname(excel_path), "Outstanding.xlsx")
    processed_rows.to_excel(processed_output_path, index=False, engine='openpyxl')

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python script_name.py <path_to_excel_file>")
        sys.exit(1)

    excel_path = sys.argv[1]
    process_excel(excel_path)
