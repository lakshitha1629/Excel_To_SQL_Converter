# Excel to SQL Converter

Convert Excel files to SQL `CREATE TABLE` and `INSERT INTO` statements.

## Installation

1. **Install the required libraries**
   
   Clone this repository to your local machine:
   ```
   pip install -r requirements.txt
   ```
2. **To run the script directly**
   ```
   python script.py
   ```
3. **To convert the Python script into an executable**
   ```
   python -m PyInstaller script.py
   ```
   This will generate an executable file in the dist directory.
4. **To read an Excel file and convert it to SQL statements, you can use the provided run.bat file or execute the following command**
   ```
   script.exe D:\investor_users_161023.xlsx
   ```
   Change the path D:\investor_users_161023.xlsx to the path of your Excel file.

***Notes***
Ensure the provided Excel file follows the expected format.
The table name will be generated with a "d" prefix based on the Excel file name.
