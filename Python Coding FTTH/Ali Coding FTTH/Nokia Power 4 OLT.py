#لاستخراج بورات الاولتيات النوكيا حسب الايبيات المعطاة في الشيت


import pandas as pd
import paramiko
import time
from concurrent.futures import ThreadPoolExecutor
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Read Excel file using pandas
df = pd.read_excel(r'IP Nokia.xlsx')

# Create an empty DataFrame to store the data
all_data = pd.DataFrame()

# Define the handle_session function
def handle_session(hostname, username, password, port):

    max_retries = 3  # Maximum number of SSH connection retries
    retry_delay = 5  # Delay between SSH connection retries (in seconds)

    for attempt in range(max_retries):
        try:
            # Create SSH client object
            ssh = paramiko.SSHClient()
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            # Establish SSH connection
            ssh.connect(hostname=hostname, username=username, password=password, port=port, timeout=120)

            # create shell object to interact with device
            shell = ssh.invoke_shell()

            # send command to device to made idle-timeout 1450 sec
            shell.send("configure system security ssh server-profile idle-timeout 1450\n")
            time.sleep(1)
            ssh.close()
            # Create SSH client object

            ssh = paramiko.SSHClient()
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            # Establish SSH connection
            ssh.connect(hostname=hostname, username=username, password=password, port=port, timeout=120)

            # create shell object to interact with device
            shell = ssh.invoke_shell()

            # wait for command output to be generated
            time.sleep(1)
            shell.send("show equipment ont interface\n")
            time.sleep(25)
            # send command to device to print users optics
            shell.send("show equipment ont optics\n")
            # wait for command output to be generated
            time.sleep(1000)

            shell.send("configure system security ssh server-profile idle-timeout 60\n")

            # Receive the output from the device
            output = b''
            while shell.recv_ready():
                buffer = shell.recv(1024)
                output += buffer

            output = output.decode('latin-1')  # Decode using 'latin-1' encoding

            # Parse the output and save it to the DataFrame
            lines = output.splitlines()
            data = [line.split() for line in lines]
            device_data = pd.DataFrame(data)

            # Append the device data to the overall data
            global all_data
            all_data = pd.concat([all_data, device_data], ignore_index=True)

            # Close SSH connection
            ssh.close()

            # Break the retry loop if the SSH connection and data processing are successful
            break

        except paramiko.AuthenticationException as auth_exception:
            print(f"Authentication failed for {hostname}: {str(auth_exception)}")
            break  # Break the retry loop on authentication failure
        except paramiko.SSHException as ssh_exception:
            print(f"SSH connection failed for {hostname}: {str(ssh_exception)}")
        except Exception as e:
            print(f"An error occurred for {hostname}: {str(e)}")

        # Close SSH connection if it's still open
        if ssh:
            ssh.close()

        if attempt < max_retries - 1:
            print(f"Retrying SSH connection to {hostname} in {retry_delay} seconds...")
            time.sleep(retry_delay)
        else:
            print(f"Max retries exceeded. Could not establish SSH connection to {hostname}.")

# Create a thread pool with a maximum of 8 threads
pool = ThreadPoolExecutor(max_workers=1)

# Submit tasks to the thread pool for each device
for index, row in df.iterrows():
    hostname = row['hostname']
    username = row['username']
    password = row['password']
    port = row['port']

    pool.submit(handle_session, hostname, username, password, port)

# Shutdown the thread pool to wait for all tasks to complete
pool.shutdown()

# Write the combined data to an Excel file
all_data.to_excel('Test.xlsx', index=False)