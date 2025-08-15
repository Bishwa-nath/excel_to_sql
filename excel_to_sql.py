import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import pandas as pd
import os

# Convert Python value to SQL-safe string
def sql_value(val):
    if pd.isna(val):
        return "NULL"
    elif str(val).lower() == 'true':
        return 1
    elif str(val).lower() == 'false':
        return 0
    elif isinstance(val, str):
        return "'" + val.replace("'", "''") + "'"
    elif isinstance(val, pd.Timestamp):
        return f"'{val.strftime('%Y-%m-%d %H:%M:%S')}'"
    else:
        return str(int(val)) if isinstance(val, float) and val.is_integer() else str(val)

# Main function to generate SQL and save directly
def generate_sql():
    file_path = file_entry.get()
    model_name = model_entry.get().strip()
    use_identity = identity_var.get()

    if not os.path.isfile(file_path):
        messagebox.showerror("Error", "Invalid file path.")
        return

    if not model_name:
        messagebox.showerror("Error", "Please enter a model/table name.")
        return

    try:
        if file_path.endswith(".csv"):
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read file:\n{str(e)}")
        return

    columns = df.columns.tolist()
    column_list_str = ", ".join(columns)

    insert_statements = []
    for _, row in df.iterrows():
        values = [sql_value(row[col]) for col in columns]
        values_str = ", ".join(values)
        insert_sql = f"INSERT INTO {model_name} ({column_list_str}) VALUES ({values_str});"
        insert_statements.append(insert_sql)

    # Add IDENTITY_INSERT if checkbox is selected
    if use_identity:
        insert_statements.insert(0, f"SET IDENTITY_INSERT {model_name} ON;")
        insert_statements.append(f"SET IDENTITY_INSERT {model_name} OFF;")

    # Save SQL file directly
    output_dir = os.path.dirname(file_path)
    output_filename = f"{model_name}_Insert.sql"
    output_path = os.path.join(output_dir, output_filename)

    try:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write("\n".join(insert_statements))
        messagebox.showinfo("Success", f"SQL script saved to:\n{output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save SQL file:\n{str(e)}")

# GUI Layout
root = tk.Tk()
root.title("Excel/CSV to SQL Generator")
root.geometry("500x260")

# File input
ttk.Label(root, text="Select Excel/CSV File:").pack(pady=(10, 0))
file_frame = tk.Frame(root)
file_frame.pack(pady=5)
file_entry = tk.Entry(file_frame, width=50)
file_entry.pack(side=tk.LEFT, padx=(0, 5))

def browse_file():
    path = filedialog.askopenfilename(filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv")])
    if path:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, path)

ttk.Button(file_frame, text="Browse", command=browse_file).pack(side=tk.LEFT)

# Model name input
ttk.Label(root, text="Enter Table Name:").pack(pady=(10, 0))
model_entry = tk.Entry(root, width=50)
model_entry.pack(pady=5)

# Identity insert checkbox
identity_var = tk.BooleanVar()
ttk.Checkbutton(root, text="Use SET IDENTITY_INSERT ON/OFF", variable=identity_var).pack(pady=5)

# Generate button
ttk.Button(root, text="Generate SQL File", command=generate_sql).pack(pady=15)

# Run app
root.mainloop()
