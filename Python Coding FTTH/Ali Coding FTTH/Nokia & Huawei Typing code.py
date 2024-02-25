#لطباعة كود داخل مجموعة من الاولتيات

import pandas as pd
import paramiko
import time

# Read the Excel sheet into a pandas DataFrame
df = pd.read_excel('N.xlsx')

# Create an empty DataFrame to store the data
all_data = pd.DataFrame()

# Iterate over each row in the dataframe
for index, row in df.iterrows():
    try:

    # Retrieve device information from each row
        hostname = row['hostname']
        username = row['username']
        password = row['password']
        port = row['port']


        # Create SSH client object
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

        # Establish SSH connection
        ssh.connect(hostname=hostname, username=username, password=password, port=port, timeout=120)

        # create shell object to interact with device
        shell = ssh.invoke_shell()

        # send commands to device
        shell.send("configure system security filetransfer protocol sftp \n")
        shell.send("admin save \n")
        
        time.sleep(5)
        
        
        # Close SSH connection
        ssh.close()
        
    except Exception as e:
        print(f"Error occurred for row {index + 1}: {str(e)}")
