import cx_Oracle
import pandas as pd
import time
import os
from tkinter import Tk, Label, Entry, Button, filedialog, StringVar, OptionMenu, messagebox, DoubleVar, X
from tkinter import ttk
# Oracle database connection details
# def execute_oracle_query(lib_dir, progress_var):
print("start")
oracle_host = 'host name'
oracle_port = 'port number'
oracle_sid = 'sid name'
oracle_user = 'username'
oracle_password = 'password'

# Shared network location to save results
shared_network_location = r'\\path\to\shared\location'

# Oracle query to be executed
oracle_query = """
select SITE_NAME, LATITUDE,LONGITUDE,BAND_INFO,BAND_CLASS_NAME,TECHNOLOGY,PCI,PSS_ID,SSS_ID,PRACH_ROOT_SEQUENCES from SITES_ATOLL_V where ATOLL_SCHEMA in ('carolinas','nashville') AND SITE_VERSION in ('0000','0001')
"""

# Directory where Oracle Instant Client is installed.
# Update this path based on your Oracle Instant Client installation.
lib_dir = r"C:\ORACLE\instantclient_18_3"

try:
    cx_Oracle.init_oracle_client(lib_dir=lib_dir)
    # print("success")
except Exception as err:
    # print("Error connecting: cx_Oracle.init_oracle_client()")
    print(err)
    # messagebox.showerror("Error",f"An error occurred: {str(err)}")

    


    # Establish connection to Oracle database
    # try:   
dsn = cx_Oracle.makedsn(host=oracle_host, port=oracle_port , sid=oracle_sid )    

connection = cx_Oracle.connect(
    user=oracle_user,
    password=oracle_password,
    # dsn=f'{oracle_host}:{oracle_port}/{oracle_sid}'
    dsn= dsn
)

start_time = time.time()
# Execute the query
cursor = connection.cursor()
cursor.execute(oracle_query)

# Fetch all results
results = cursor.fetchall()

# Close cursor and connection


# Convert results to pandas DataFrame
df = pd.DataFrame(results, columns=[col[0] for col in cursor.description])

# Save DataFrame to Excel file in shared network location
excel_filename = 'oracle_query_results.xlsx'
excel_file_path = f'{shared_network_location}/{excel_filename}'

if os.path.exists(excel_file_path):
    os.remove(excel_file_path)

df.to_excel(excel_file_path, index=False)

# messagebox.showinfo(f"Query results saved to: {excel_file_path}")

cursor.close()
connection.close()
    # except Exception as err:
    #      messagebox.showerror("Error",f"An error occurred: {str(err)}")
    # finally:
    #      progress_var.set(100)

# window = Tk()
# window.title("Oracle Query Executor")

# lib_dir_label = Label(window, text="Client Directory:")
# lib_dir_label.pack(pady=(20,5))
# lib_dir_entry = Entry(window, width = 70)

# lib_dir_entry.pack(pady= 5)

# progress_var = DoubleVar()
# progress_bar = ttk.Progressbar(window, variable=progress_var, maximum=100)
# progress_bar.pack(fill=X, padx=10,pady=5)

# def execute_query():
#      lib_dir = lib_dir_entry.get()
#      progress_var.set(0)
#      execute_oracle_query(lib_dir, progress_var)

# button = Button(window, text="Execute Oracle Query", command= execute_query)
# button.pack()

# window.mainloop()
         
    # print("The execution time was..{}".format(time.time()-start_time))
# print("done")


