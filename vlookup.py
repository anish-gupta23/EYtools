import tkinter as tk
import numpy as np
from tkinter import filedialog, messagebox
import pandas as pd

def merge_and_vlookup(file1, file2, output_file1, output_file2):
    try:
        # Read Excel files with specified date format
        df1 = pd.read_excel(file1, parse_dates=['DocumentDate'], date_format='%Y-%m-%d')
        df2 = pd.read_excel(file2, parse_dates=['DocumentDate'], date_format='%Y-%m-%d')

        df1['KEY'] = df1['SupplierGSTIN'].astype(str) + df1['DocumentType'].astype(str) + df1['DocumentNumber'].astype(str)
        df2['KEY'] = df2['SupplierGSTIN'].astype(str) + df2['DocumentType'].astype(str) + df2['DocumentNumber'].astype(str)

        merged_df1 = pd.merge(df1, df2, on='KEY', how='left', suffixes=('_F1', '_F2'))
        merged_df2 = pd.merge(df2, df1, on='KEY', how='left', suffixes=('_F2', '_F1'))

        merged_df1.fillna("NA", inplace=True)
        merged_df2.fillna("NA", inplace=True)

        file1_cols_1 = [col for col in merged_df1.columns if col.endswith('_F1')]
        file1_cols_2 = [col for col in merged_df2.columns if col.endswith('_F1')]

        merged_df1[file1_cols_1] = merged_df1[file1_cols_1].where(~merged_df1[file1_cols_1].eq('NA').all(axis=1), other="NA")
        merged_df2[file1_cols_2] = merged_df2[file1_cols_2].where(~merged_df2[file1_cols_2].eq('NA').all(axis=1), other="NA")

        file2_cols_1 = [col for col in merged_df1.columns if col.endswith('_F2')]
        file2_cols_2 = [col for col in merged_df2.columns if col.endswith('_F2')]

        merged_df1[file2_cols_1] = merged_df1[file2_cols_1].where(~merged_df1[file2_cols_1].eq('NA').all(axis=1), other="NA")
        merged_df2[file2_cols_2] = merged_df2[file2_cols_2].where(~merged_df2[file2_cols_2].eq('NA').all(axis=1), other="NA")


        integer_columns = ['TaxableValue_F1', 'TaxableValue_F2', 'IntegratedTaxAmount_F1', 'IntegratedTaxAmount_F2',
                        'CentralTaxAmount_F1', 'CentralTaxAmount_F2', 'StateUTTaxAmount_F1', 'StateUTTaxAmount_F2']

        for col in integer_columns:
            merged_df1[col] = pd.to_numeric(merged_df1[col], errors='coerce')
            merged_df2[col] = pd.to_numeric(merged_df2[col], errors='coerce')

        for col in integer_columns:
            merged_df1[col] = merged_df1[col].apply(lambda x: "NA" if pd.isnull(x) else x)
            merged_df2[col] = merged_df2[col].apply(lambda x: "NA" if pd.isnull(x) else x)

        merged_df1['TaxableValue_F1'] = pd.to_numeric(merged_df1['TaxableValue_F1'], errors='coerce')
        merged_df1['TaxableValue_F2'] = pd.to_numeric(merged_df1['TaxableValue_F2'], errors='coerce')
        merged_df2['TaxableValue_F1'] = pd.to_numeric(merged_df2['TaxableValue_F1'], errors='coerce')
        merged_df2['TaxableValue_F2'] = pd.to_numeric(merged_df2['TaxableValue_F2'], errors='coerce')

        merged_df1['TaxableValue_difference'] = merged_df1['TaxableValue_F1'] - merged_df1['TaxableValue_F2']
        merged_df2['TaxableValue_difference'] = merged_df2['TaxableValue_F2'] - merged_df2['TaxableValue_F1']

        merged_df1['TaxableValue_difference'].fillna("NA", inplace=True)
        merged_df2['TaxableValue_difference'].fillna("NA", inplace=True)

        merged_df1['IntegratedTaxAmount_F1'] = pd.to_numeric(merged_df1['IntegratedTaxAmount_F1'], errors='coerce')
        merged_df1['IntegratedTaxAmount_F2'] = pd.to_numeric(merged_df1['IntegratedTaxAmount_F2'], errors='coerce')
        merged_df2['IntegratedTaxAmount_F1'] = pd.to_numeric(merged_df2['IntegratedTaxAmount_F1'], errors='coerce')
        merged_df2['IntegratedTaxAmount_F2'] = pd.to_numeric(merged_df2['IntegratedTaxAmount_F2'], errors='coerce')

        merged_df1['IntegratedTaxAmount_difference'] = merged_df1['IntegratedTaxAmount_F1'] - merged_df1['IntegratedTaxAmount_F2']
        merged_df2['IntegratedTaxAmount_difference'] = merged_df2['IntegratedTaxAmount_F2'] - merged_df2['IntegratedTaxAmount_F1']

        merged_df1['IntegratedTaxAmount_difference'].fillna("NA", inplace=True)
        merged_df2['IntegratedTaxAmount_difference'].fillna("NA", inplace=True)

        merged_df1['CentralTaxAmount_F1'] = pd.to_numeric(merged_df1['CentralTaxAmount_F1'], errors='coerce')
        merged_df1['CentralTaxAmount_F2'] = pd.to_numeric(merged_df1['CentralTaxAmount_F2'], errors='coerce')
        merged_df2['CentralTaxAmount_F1'] = pd.to_numeric(merged_df2['CentralTaxAmount_F1'], errors='coerce')
        merged_df2['CentralTaxAmount_F2'] = pd.to_numeric(merged_df2['CentralTaxAmount_F2'], errors='coerce')

        merged_df1['CentralTaxAmount_difference'] = merged_df1['CentralTaxAmount_F1'] - merged_df1['CentralTaxAmount_F2']
        merged_df2['CentralTaxAmount_difference'] = merged_df2['CentralTaxAmount_F2'] - merged_df2['CentralTaxAmount_F1']

        merged_df1['CentralTaxAmount_difference'].fillna("NA", inplace=True)
        merged_df2['CentralTaxAmount_difference'].fillna("NA", inplace=True)

        merged_df1['StateUTTaxAmount_F1'] = pd.to_numeric(merged_df1['StateUTTaxAmount_F1'], errors='coerce')
        merged_df1['StateUTTaxAmount_F2'] = pd.to_numeric(merged_df1['StateUTTaxAmount_F2'], errors='coerce')
        merged_df2['StateUTTaxAmount_F1'] = pd.to_numeric(merged_df2['StateUTTaxAmount_F1'], errors='coerce')
        merged_df2['StateUTTaxAmount_F2'] = pd.to_numeric(merged_df2['StateUTTaxAmount_F2'], errors='coerce')

        merged_df1['StateUTTaxAmount_difference'] = merged_df1['StateUTTaxAmount_F1'] - merged_df1['StateUTTaxAmount_F2']
        merged_df2['StateUTTaxAmount_difference'] = merged_df2['StateUTTaxAmount_F2'] - merged_df2['StateUTTaxAmount_F1']

        merged_df1['StateUTTaxAmount_difference'].fillna("NA", inplace=True)
        merged_df2['StateUTTaxAmount_difference'].fillna("NA", inplace=True)

        merged_df1['DocumentDate_difference'] = merged_df1.apply(lambda row: 'NA' if pd.isnull(row['DocumentDate_F1']) or pd.isnull(row['DocumentDate_F2']) else 
                                                (row['DocumentDate_F1'] == row['DocumentDate_F2']), axis=1).astype(bool)

        merged_df2['DocumentDate_difference'] = merged_df2.apply(lambda row: 'NA' if pd.isnull(row['DocumentDate_F1']) or pd.isnull(row['DocumentDate_F2']) else 
                                                (row['DocumentDate_F1'] == row['DocumentDate_F2']), axis=1).astype(bool)

        merged_df1['CustomerGSTIN_F1'].fillna("NA", inplace=True)
        merged_df1['CustomerGSTIN_F2'].fillna("NA", inplace=True)
        merged_df2['CustomerGSTIN_F1'].fillna("NA", inplace=True)
        merged_df2['CustomerGSTIN_F2'].fillna("NA", inplace=True)

        merged_df1['CustomerGSTIN_difference'] = merged_df1.apply(lambda row: 'TRUE' if row['CustomerGSTIN_F1'] == row['CustomerGSTIN_F2'] else 
                                                                ('FALSE' if row['CustomerGSTIN_F1'] != row['CustomerGSTIN_F2'] else 'NA'), axis=1)

        merged_df2['CustomerGSTIN_difference'] = merged_df2.apply(lambda row: 'TRUE' if row['CustomerGSTIN_F1'] == row['CustomerGSTIN_F2'] else 
                                                                ('FALSE' if row['CustomerGSTIN_F1'] != row['CustomerGSTIN_F2'] else 'NA'), axis=1)

        merged_df1['TaxableValue_difference'].replace('NA', 0, inplace=True)
        merged_df2['TaxableValue_difference'].replace('NA', 0, inplace=True)
        merged_df1['IntegratedTaxAmount_difference'].replace('NA', 0, inplace=True)
        merged_df2['IntegratedTaxAmount_difference'].replace('NA', 0, inplace=True)
        merged_df1['CentralTaxAmount_difference'].replace('NA', 0, inplace=True)
        merged_df2['CentralTaxAmount_difference'].replace('NA', 0, inplace=True)
        merged_df1['StateUTTaxAmount_difference'].replace('NA', 0, inplace=True)
        merged_df2['StateUTTaxAmount_difference'].replace('NA', 0, inplace=True)

        merged_df1['TaxableValue_difference'] = pd.to_numeric(merged_df1['TaxableValue_difference'])
        merged_df2['TaxableValue_difference'] = pd.to_numeric(merged_df2['TaxableValue_difference'])
        merged_df1['IntegratedTaxAmount_difference'] = pd.to_numeric(merged_df1['IntegratedTaxAmount_difference'])
        merged_df2['IntegratedTaxAmount_difference'] = pd.to_numeric(merged_df2['IntegratedTaxAmount_difference'])
        merged_df1['CentralTaxAmount_difference'] = pd.to_numeric(merged_df1['CentralTaxAmount_difference'])
        merged_df2['CentralTaxAmount_difference'] = pd.to_numeric(merged_df2['CentralTaxAmount_difference'])
        merged_df1['StateUTTaxAmount_difference'] = pd.to_numeric(merged_df1['StateUTTaxAmount_difference'])
        merged_df2['StateUTTaxAmount_difference'] = pd.to_numeric(merged_df2['StateUTTaxAmount_difference'])


        merged_df1['mismatched_reason'] = ''
        merged_df1.loc[merged_df1['CustomerGSTIN_difference'] == 'False', 'mismatched_reason'] += 'CustomerGSTIN/'
        merged_df1.loc[~merged_df1['DocumentDate_difference'], 'mismatched_reason'] += 'DocumentDate/'
        merged_df1.loc[((merged_df1['TaxableValue_difference'] < -10) | (merged_df1['TaxableValue_difference'] > 10)) & (~merged_df1['TaxableValue_difference'].isna()), 'mismatched_reason'] += 'TaxableValue/'
        merged_df1.loc[((merged_df1['IntegratedTaxAmount_difference'] < -10) | (merged_df1['IntegratedTaxAmount_difference'] > 10)) & (~merged_df1['IntegratedTaxAmount_difference'].isna()), 'mismatched_reason'] += 'IGST/'
        merged_df1.loc[((merged_df1['CentralTaxAmount_difference'] < -10) | (merged_df1['CentralTaxAmount_difference'] > 10)) & (~merged_df1['CentralTaxAmount_difference'].isna()), 'mismatched_reason'] += 'CGST/'
        merged_df1.loc[((merged_df1['StateUTTaxAmount_difference'] < -10) | (merged_df1['StateUTTaxAmount_difference'] > 10)) & (~merged_df1['StateUTTaxAmount_difference'].isna()), 'mismatched_reason'] += 'SGST/'

        merged_df1.loc[merged_df1.apply(lambda row: 'Additional Entries' if 'NA' in row.values else '', axis=1) != '', 'mismatched_reason'] = 'Additional Entries'
        merged_df1.loc[merged_df1['mismatched_reason'] == '', 'mismatched_reason'] = 'Exact Match'

        merged_df2['mismatched_reason'] = ''
        merged_df2.loc[merged_df2['CustomerGSTIN_difference'] == 'False', 'mismatched_reason'] += 'CustomerGSTIN/'
        merged_df2.loc[~merged_df2['DocumentDate_difference'], 'mismatched_reason'] += 'DocumentDate/'
        merged_df2.loc[((merged_df2['TaxableValue_difference'] < -10) | (merged_df2['TaxableValue_difference'] > 10)) & (~merged_df2['TaxableValue_difference'].isna()), 'mismatched_reason'] += 'TaxableValue/'
        merged_df2.loc[((merged_df2['IntegratedTaxAmount_difference'] < -10) | (merged_df2['IntegratedTaxAmount_difference'] > 10)) & (~merged_df2['IntegratedTaxAmount_difference'].isna()), 'mismatched_reason'] += 'IGST/'
        merged_df2.loc[((merged_df2['CentralTaxAmount_difference'] < -10) | (merged_df2['CentralTaxAmount_difference'] > 10)) & (~merged_df2['CentralTaxAmount_difference'].isna()), 'mismatched_reason'] += 'CGST/'
        merged_df2.loc[((merged_df2['StateUTTaxAmount_difference'] < -10) | (merged_df2['StateUTTaxAmount_difference'] > 10)) & (~merged_df2['StateUTTaxAmount_difference'].isna()), 'mismatched_reason'] += 'SGST/'

        merged_df2.loc[merged_df2.apply(lambda row: 'Additional Entries' if 'NA' in row.values else '', axis=1) != '', 'mismatched_reason'] = 'Additional Entries'
        merged_df2.loc[merged_df2['mismatched_reason'] == '', 'mismatched_reason'] = 'Exact Match'

        merged_df1['Report Type'] = ''
        merged_df1.loc[merged_df1['mismatched_reason'].str.contains('CustomerGSTIN'), 'Report Type'] = 'CustomerGSTIN Error'
        merged_df1.loc[merged_df1['mismatched_reason'].str.contains('DocumentDate'), 'Report Type'] = 'DocumentDate Mismatch'
        merged_df1.loc[merged_df1['mismatched_reason'].str.contains('CGST|SGST|IGST|TaxableValue'), 'Report Type'] = 'Value Mismatch'
        merged_df1.loc[merged_df1['mismatched_reason'].str.count('/') >= 2, 'Report Type'] = 'Multi Mismatch'
        merged_df1.loc[merged_df1['mismatched_reason'] == 'Exact Match', 'Report Type'] = 'Exact Match'
        merged_df1.loc[merged_df1['mismatched_reason'] == 'Additional Entries', 'Report Type'] = 'Additional Entries'

        merged_df2['Report Type'] = ''
        merged_df2.loc[merged_df2['mismatched_reason'].str.contains('CustomerGSTIN'), 'Report Type'] = 'CustomerGSTIN Error'
        merged_df2.loc[merged_df2['mismatched_reason'].str.contains('DocumentDate'), 'Report Type'] = 'DocumentDate Mismatch'
        merged_df2.loc[merged_df2['mismatched_reason'].str.contains('CGST|SGST|IGST|TaxableValue'), 'Report Type'] = 'Value Mismatch'
        merged_df2.loc[merged_df2['mismatched_reason'].str.count('/') >= 2, 'Report Type'] = 'Multi Mismatch'
        merged_df2.loc[merged_df2['mismatched_reason'] == 'Exact Match', 'Report Type'] = 'Exact Match'
        merged_df2.loc[merged_df2['mismatched_reason'] == 'Additional Entries', 'Report Type'] = 'Additional Entires'


        merged_df1.to_excel(output_file1, index=False)
        merged_df2.to_excel(output_file2, index=False)

        report_type_counts_1 = merged_df1['Report Type'].value_counts()
        report_counts_df_1 = pd.DataFrame(report_type_counts_1.items(), columns=['Report Type', 'Count'])

        report_type_counts_2 = merged_df2['Report Type'].value_counts()
        report_counts_df_2 = pd.DataFrame(report_type_counts_2.items(), columns=['Report Type', 'Count'])

        sum_values_df1 = merged_df1.groupby('Report Type').agg({
            'TaxableValue_F1': 'sum',
            'IntegratedTaxAmount_F1': 'sum',
            'CentralTaxAmount_F1': 'sum',
            'StateUTTaxAmount_F1': 'sum'
        }).reset_index()

        # Calculate sum of values before lookup for each mismatch type in sheet 2
        sum_values_df2 = merged_df2.groupby('Report Type').agg({
            'TaxableValue_F2': 'sum',
            'IntegratedTaxAmount_F2': 'sum',
            'CentralTaxAmount_F2': 'sum',
            'StateUTTaxAmount_F2': 'sum'
        }).reset_index()

        # Merge count data with sum values for sheet 1
        report_counts_df_1 = pd.merge(report_counts_df_1, sum_values_df1, on='Report Type', how='left')

        # Merge count data with sum values for sheet 2
        report_counts_df_2 = pd.merge(report_counts_df_2, sum_values_df2, on='Report Type', how='left')

        with pd.ExcelWriter(output_file1, mode='a', engine='openpyxl') as writer1:
            report_counts_df_1.to_excel(writer1, sheet_name='Report_Counts_1', index=False, startcol=2)

        with pd.ExcelWriter(output_file2, mode='a', engine='openpyxl') as writer2:
            report_counts_df_2.to_excel(writer2, sheet_name='Report_Counts_2', index=False, startcol=2)

        messagebox.showinfo("Success", "Merging and VLOOKUP completed successfully!")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def browse_files():
    file1_path = filedialog.askopenfilename(title="Select Excel file 1")
    if file1_path:
        input_entry1.delete(0, tk.END)
        input_entry1.insert(0, file1_path)
        
    file2_path = filedialog.askopenfilename(title="Select Excel file 2")
    if file2_path:
        input_entry2.delete(0, tk.END)
        input_entry2.insert(0, file2_path)
    if file1_path and file2_path:
        output_file1_path = filedialog.asksaveasfilename(title="Save Merged File 1 As", defaultextension=".xlsx")
        output_file2_path = filedialog.asksaveasfilename(title="Save Merged File 2 As", defaultextension=".xlsx")
        if output_file1_path and output_file2_path:
            merge_and_vlookup(file1_path, file2_path, output_file1_path, output_file2_path)

root = tk.Tk()
root.title("Excel File Merger and VLOOKUP")
label1 = tk.Label(root, text="Excel File 1:")
label1.pack(pady=5)
input_entry1 = tk.Entry(root, width=50)
input_entry1.pack(pady=5)

label2 = tk.Label(root, text="Excel File 2:")
label2.pack(pady=5)
input_entry2 = tk.Entry(root, width=50)
input_entry2.pack(pady=5)

button = tk.Button(root, text="Browse Excel Files", command=browse_files)
button.pack(pady=20)

root.mainloop()
