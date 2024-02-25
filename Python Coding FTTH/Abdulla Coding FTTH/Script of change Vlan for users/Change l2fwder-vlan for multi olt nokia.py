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

        # Send command to device to set idle-timeout to 1000 sec
        shell.send("configure system security ssh server-profile idle-timeout 1000\n")
        time.sleep(5)

        # Close SSH connection
        ssh.close()

        # Establish SSH connection to the device
        ssh.connect(hostname=hostname, username=username, password=password, port=port)

        # Create shell object to interact with the device
        shell = ssh.invoke_shell()

        # Send command to device to set idle-timeout to 60 sec
        shell.send("configure system security ssh server-profile idle-timeout 60\n")
        time.sleep(2)
        shell.send("exit all\n")

        # Read the VLAN information for the specific OLT from the corresponding sheet
        vlan_df = pd.read_excel('C-Vlan.xlsx', sheet_name=olt_name)

        # Loop through the VLAN DataFrame rows to perform actions on the device
        for _, row in vlan_df.iterrows():
            port = row['port']
            vlan_id = row['vlan-id']
            inner_vlan = row['inner-vlan']
            main_vlan = row['Main-vlan']

            # Wait for the shell to process previous commands (if needed)
            time.sleep(0.5)

            # Send commands to the device using the extracted data
            shell.send(f"configure bridge {port} {vlan_id} no l2fwder-vlan \n")
            shell.send(f"configure bridge {port} {vlan_id} l2fwder-vlan stacked:{main_vlan}:{inner_vlan}\n")
            shell.send("exit all\n")
            time.sleep(0.5)

        # Wait for the commands to be processed and the output (optional)
        time.sleep(1)

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
devices_df = pd.read_excel('C-Vlan.xlsx', sheet_name='Devices')

# Create a thread pool with a maximum of 2 threads
pool = ThreadPoolExecutor(max_workers=2)

# Loop through the DataFrame rows to open sessions for each device
for _, device_row in devices_df.iterrows():
    hostname = device_row['hostname']
    username = device_row['username']
    password = device_row['password']
    port = device_row['port']
    olt_name = device_row['OLT Name']

    # Submit each device task to the thread pool
    pool.submit(handle_session, hostname, username, password, port, olt_name)

# Wait for all tasks to complete
pool.shutdown(wait=True)