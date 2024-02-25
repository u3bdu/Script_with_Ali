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
        shell.send('enable \n')
        shell.send('config \n')
        shell.send('scroll \n')
        shell.send(' \n')
        time.sleep(1)

        # Read the VLAN information for the specific OLT from the corresponding sheet
        vlan_df = pd.read_excel('C-Vlan.xlsx', sheet_name=olt_name)

        # Loop through the VLAN DataFrame rows to perform actions on the device
        for _, row in vlan_df.iterrows():
            service_port = row['Service-port-id']
            service_port_num= row['S/N']
            main_vlan = row['Main-vlan']
            Slot_Gpon = row['Slot-Gpon']
            ont_id = row['ont-id']
            inner_vlan = row['inner-vlan']

            # Wait for the shell to process previous commands (if needed)
            time.sleep(0.5)

            # Send commands to the device using the extracted data
            shell.send(f'undo service-port {service_port} \n')
            shell.send(f'service-port {service_port} vlan {main_vlan} gpon {Slot_Gpon} ont {ont_id} gemport 1 multi-service user-vlan 1 tag-transform translate-and-add inner-vlan {inner_vlan} inner-priority 0 \n')
            shell.send(' \n')
            time.sleep(0.75)
            print(service_port_num)

        time.sleep(0.5)
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