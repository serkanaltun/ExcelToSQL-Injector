import os
import openpyxl
import pyodbc
import tkinter as tk
from tkinter import filedialog

def inject_data():
    # Take names about server, database, and the file from the user
    file_path = file_entry.get()
    server_name = server_entry.get()
    database_name = database_entry.get()

    # Check if any of the required fields is empty
    if file_path == '' or server_name == '' or database_name == '':
        result_label.config(text="Please fill in all the required fields.")
        return

    try:
        # Open excel file
        wb = openpyxl.load_workbook(file_path)

        # Open the first page of the excel file
        sheet = wb.active

        # Read data and convert it to a table
        table_data = []
        for row in sheet.iter_rows(values_only=True):
            table_data.append(row)

        # Make a connection to SQL Server
        server = server_name
        database = database_name
        conn = pyodbc.connect('Driver={SQL Server};Server=' + server + ';Database=' + database + ';Trusted_Connection=yes;')

        # Transfer the data to SQL Server.
        cursor = conn.cursor()
        table_name = os.path.splitext(os.path.basename(file_path))[0]

         # Check if the table exists
        if not cursor.tables(table=table_name).fetchone():
            # Create table query
            create_table_query = f"CREATE TABLE {table_name} ("

            # Define the table columns.
            columns = table_data[0]
            for column in columns:
                create_table_query += f"{column} VARCHAR(MAX),"

            create_table_query = create_table_query.rstrip(',') + ")"

            # Create table
            cursor.execute(create_table_query)

        # Insert the data into the table.
        for data_row in table_data[1:]:
            insert_query = f"INSERT INTO {table_name} VALUES ("
            for value in data_row:
                insert_query += f"'{value}',"
            insert_query = insert_query.rstrip(',') + ")"
            cursor.execute(insert_query)

        # Save changes
        conn.commit()

        # Close the connection
        conn.close()

        # Close the excel file
        wb.close()

        result_label.config(text="All data have been successfully injected into SQL Server.")

    except Exception as e:
        result_label.config(text=f"Error: {str(e)}")

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    file_entry.delete(0, tk.END)
    file_entry.insert(0, file_path)

root = tk.Tk()
root.title("Data Injection")
root.geometry("400x300")
root.configure(bg="#F5F5F5")

header_label = tk.Label(root, text="Data Injection", font=("Arial", 16, "bold"), bg="#F5F5F5")
header_label.pack(pady=10)

file_frame = tk.Frame(root, bg="#F5F5F5")
file_frame.pack(pady=10)

file_label = tk.Label(file_frame, text="Excel file path:", bg="#F5F5F5")
file_label.pack(side=tk.LEFT)

file_entry = tk.Entry(file_frame, width=30)
file_entry.pack(side=tk.LEFT, padx=10)

browse_button = tk.Button(file_frame, text="Browse", command=browse_file)
browse_button.pack(side=tk.LEFT)

server_frame = tk.Frame(root, bg="#F5F5F5")
server_frame.pack(pady=10)

server_label = tk.Label(server_frame, text="Server name:", bg="#F5F5F5")
server_label.pack(side=tk.LEFT)

server_entry = tk.Entry(server_frame, width=30)
server_entry.pack(side=tk.LEFT, padx=10)

database_frame = tk.Frame(root, bg="#F5F5F5")
database_frame.pack(pady=10)

database_label = tk.Label(database_frame, text="Database name:", bg="#F5F5F5")
database_label.pack(side=tk.LEFT)

database_entry = tk.Entry(database_frame, width=30)
database_entry.pack(side=tk.LEFT, padx=10)

inject_button = tk.Button(root, text="Inject Data", command=inject_data, bg="#4CAF50", fg="white", font=("Arial", 12, "bold"))
inject_button.pack(pady=20)

result_label = tk.Label(root, text="", font=("Arial", 12), bg="#F5F5F5")
result_label.pack()

root.mainloop()
