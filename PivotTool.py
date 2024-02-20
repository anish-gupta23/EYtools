import tkinter as tk
from tkinter import filedialog
import pandas as pd

def create_pivot_table(input_file, rows, values):
    df = pd.read_excel(input_file)
    print("Column Names:", df.columns)  # Add this line for debugging
    
    pivot_table = pd.pivot_table(df, 
                                  values=values, 
                                  index=rows)
    
    return pivot_table

def browse_file(entry):
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if filename:
        entry.delete(0, tk.END)
        entry.insert(0, filename)

def generate_pivot():
    input_file1 = input_entry1.get()
    input_file2 = input_entry2.get()
    rows_str1 = rows_entry1.get()
    rows_str2 = rows_entry2.get()
    values_str1 = values_entry1.get()
    values_str2 = values_entry2.get()
    
    if not input_file1 or not input_file2 or not rows_str1 or not rows_str2 or not values_str1 or not values_str2:
        output_text.config(state=tk.NORMAL)
        output_text.delete('1.0', tk.END)
        output_text.insert(tk.END, "Please fill in all required fields.")
        output_text.config(state=tk.DISABLED)
        return
    
    rows1 = [row.strip() for row in rows_str1.split(',')]
    rows2 = [row.strip() for row in rows_str2.split(',')]
    values1 = [val.strip() for val in values_str1.split(',')]
    values2 = [val.strip() for val in values_str2.split(',')]
    
    pivot_table1 = create_pivot_table(input_file1, rows1, values1)
    pivot_table2 = create_pivot_table(input_file2, rows2, values2)
    
    save_path1 = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not save_path1:
        return
    
    save_path2 = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not save_path2:
        return
    
    pivot_table1.to_excel(save_path1, merge_cells=False)
    pivot_table2.to_excel(save_path2, merge_cells=False)
    
    output_text.config(state=tk.NORMAL)
    output_text.delete('1.0', tk.END)
    output_text.insert(tk.END, f"Pivot tables saved to {save_path1} and {save_path2}")
    output_text.config(state=tk.DISABLED)

root = tk.Tk()
root.title("Excel Pivot Table Generator")

frame1 = tk.LabelFrame(root, text="File 1")
frame1.grid(row=0, column=0, padx=10, pady=10)

input_label1 = tk.Label(frame1, text="Input Excel File:")
input_label1.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
input_entry1 = tk.Entry(frame1, width=30)
input_entry1.grid(row=0, column=1, padx=5, pady=5)
browse_button1 = tk.Button(frame1, text="Browse", command=lambda: browse_file(input_entry1))
browse_button1.grid(row=0, column=2, padx=5, pady=5)

rows_label1 = tk.Label(frame1, text="Rows (comma-separated):")
rows_label1.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
rows_entry1 = tk.Entry(frame1, width=30)
rows_entry1.grid(row=1, column=1, padx=5, pady=5)

values_label1 = tk.Label(frame1, text="Values (comma-separated):")
values_label1.grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
values_entry1 = tk.Entry(frame1, width=30)
values_entry1.grid(row=2, column=1, padx=5, pady=5)

frame2 = tk.LabelFrame(root, text="File 2")
frame2.grid(row=0, column=1, padx=10, pady=10)

input_label2 = tk.Label(frame2, text="Input Excel File:")
input_label2.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
input_entry2 = tk.Entry(frame2, width=30)
input_entry2.grid(row=0, column=1, padx=5, pady=5)
browse_button2 = tk.Button(frame2, text="Browse", command=lambda: browse_file(input_entry2))
browse_button2.grid(row=0, column=2, padx=5, pady=5)

rows_label2 = tk.Label(frame2, text="Rows (comma-separated):")
rows_label2.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
rows_entry2 = tk.Entry(frame2, width=30)
rows_entry2.grid(row=1, column=1, padx=5, pady=5)

values_label2 = tk.Label(frame2, text="Values (comma-separated):")
values_label2.grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
values_entry2 = tk.Entry(frame2, width=30)
values_entry2.grid(row=2, column=1, padx=5, pady=5)

generate_button = tk.Button(root, text="Generate Pivot Tables", command=generate_pivot)
generate_button.grid(row=1, column=0, columnspan=2, padx=10, pady=10)

output_text = tk.Text(root, height=3, width=60)
output_text.grid(row=2, column=0, columnspan=2, padx=10, pady=10)
output_text.config(state=tk.DISABLED)

root.mainloop()

