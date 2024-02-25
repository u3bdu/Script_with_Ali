import paramiko
import time
from concurrent.futures import ThreadPoolExecutor
import pandas as pd

# Define the handle_session function to perform actions on a device
def handle_session(hostname, username, password, port, olt_name):
    try:
        # Convert olt_name to string
        olt_name = str(olt_name)

        # Create SSH client object
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

        # Establish SSH connection to the device
        ssh.connect(hostname=hostname, username=username, password=password, port=port)

        # Create shell object to interact with the device
        shell = ssh.invoke_shell()
        shell.send('configure system sntp no enable\n')
        shell.send('configure system sntp timezone-offset 180\n')
        time.sleep(1)
        shell.send(' configure system time ntp  server 37.236.221.6\n')
        time.sleep(1)
        shell.send('configure system time ntp  server 37.237.71.2\n')
        time.sleep(1)
        shell.send('configure system time ntp  server 37.239.240.6\n')
        time.sleep(1)
        shell.send('configure system time ntp no shutdown\n')
        time.sleep(1)
        shell.send('admin save\n')

        # Close SSH connection
        ssh.close()

        # Output Data
        output = olt_name
        print(output)

    except ValueError as ve:
        print(f"Invalid port value for {olt_name}: {str(ve)}")
    except paramiko.AuthenticationException as auth_exception:
        print(f"Authentication failed for {olt_name}: {str(auth_exception)}")
    except paramiko.SSHException as ssh_exception:
        print(f"SSH connection failed for {olt_name}: {str(ssh_exception)}")
    except Exception as e:
        print(f"An error occurred for {olt_name}: {str(e)}")

# Read data from Excel sheet into DataFrames
devices_df = pd.read_excel(r"D:\Earthlink\New folder\Notes\py\dataN-.xlsx", sheet_name='Nokia')

# Create a thread pool with a maximum of 20 threads
pool = ThreadPoolExecutor(max_workers=20)

# Loop through the DataFrame rows to open sessions for each device
for _, device_row in devices_df.iterrows():
    hostname = device_row['hostname']
    username = device_row['username']
    password = device_row['password']
    port = device_row['port']
    olt_name = device_row['olt name']

    # Submit each device task to the thread pool
    pool.submit(handle_session, hostname, username, password, port, olt_name)

# Wait for all tasks to complete
pool.shutdown(wait=True)