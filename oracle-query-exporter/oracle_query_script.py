import cx_Oracle
import pandas as pd
import os

# Oracle database connection details
oracle_host = 'host name'
oracle_port = 'port number'
oracle_sid = 'sid name'
oracle_user = 'username'
oracle_password = 'password'

# Shared network location to save results
shared_network_location = r'\\path\to\shared\location'

# Oracle query to be executed
oracle_query = """
SELECT SITE_NAME, LATITUDE, LONGITUDE, BAND_INFO, BAND_CLASS_NAME, TECHNOLOGY, PCI, PSS_ID, SSS_ID, PRACH_ROOT_SEQUENCES FROM SITES_ATOLL_V WHERE ATOLL_SCHEMA IN ('carolinas', 'nashville') AND SITE_VERSION IN ('0000', '0001')
"""

# Directory where Oracle Instant Client is installed.
# Update this path based on your Oracle Instant Client installation.
lib_dir = r"C:\ORACLE\instantclient_18_3"

# Initialize cx_Oracle client
cx_Oracle.init_oracle_client(lib_dir=lib_dir)

# Establish connection to Oracle database
dsn = cx_Oracle.makedsn(host=oracle_host, port=oracle_port, sid=oracle_sid)
connection = cx_Oracle.connect(
    user=oracle_user,
    password=oracle_password,
    dsn=dsn
)

# Execute the query
cursor = connection.cursor()
cursor.execute(oracle_query)

# Fetch all results
results = cursor.fetchall()

# Close cursor and connection
cursor.close()
connection.close()

# Convert results to pandas DataFrame
df = pd.DataFrame(results, columns=[col[0] for col in cursor.description])

# Save DataFrame to Excel file in shared network location
excel_filename = 'oracle_query_results.xlsx'
excel_file_path = os.path.join(shared_network_location, excel_filename)

if os.path.exists(excel_file_path):
    os.remove(excel_file_path)

df.to_excel(excel_file_path, index=False)
