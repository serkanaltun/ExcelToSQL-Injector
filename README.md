# Excel to SQL Server Table Injector

This project is a simple Python application that allows users to inject data from an Excel file into a SQL Server database. It utilizes the `openpyxl` library to read data from the Excel file and the `pyodbc` library to establish a connection with the SQL Server database and perform data injection.

## Features

- Select an Excel file using a file dialog.
- Specify the server name and database name for the SQL Server.
- Read data from the selected Excel file.
- Create a table in the SQL Server database if it doesn't exist.
- Inject the data into the table.
- Display the result of the data injection process.

## Notes

- The Excel file should have data starting from the first row.
- The first row of the Excel file will be used as column names in the SQL table.
- The table name in the SQL Server will be derived from the Excel file name.
- The data types in the SQL table will be set to `VARCHAR(MAX)` for all columns.
- The SQL Server should be accessible and allow trusted connections.

## Executable Version

If you prefer to use the pre-built executable version of this project, you can download it from the following link:
https://drive.google.com/drive/folders/1cfuVnlZk5DJdHmfmVNDs92amWDXQdB3v?usp=sharing

Feel free to modify and enhance this project according to your needs.
