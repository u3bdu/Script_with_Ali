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
        # Establish SSH connection
        ssh.connect(hostname=hostname, username=username, password=password, port=port)

        # create shell object to interact with device
        shell = ssh.invoke_shell()
            
        # wait for command output to be generated
        time.sleep(1)
            
        shell.send("configure system security ssh server-profile idle-timeout 500\n")
        time.sleep(1)
        ssh.close()

            # Create SSH client object
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        # Establish SSH connection
        ssh.connect(hostname=hostname, username=username, password=password, port=port)

        # create shell object to interact with device
        shell = ssh.invoke_shell()
            
        # wait for command output to be generated
        time.sleep(1)

        shell.send("configure system security ssh server-profile idle-timeout 60\n")
        time.sleep(1)
        shell.send('configure service vpls {vlan} v-vpls customer 1 vlan {vlan} description "HSI_vlan" \n')
        time.sleep(0.5)
        shell.send("configure service vpls {vlan} v-vpls customer 1 vlan {vlan} sap lt:1/1/1:{vlan}\n")
        time.sleep(0.5)
        shell.send("configure service vpls {vlan} v-vpls customer 1 vlan {vlan} sap lt:1/1/2:{vlan}\n")
        time.sleep(0.5)
        shell.send("configure service vpls {vlan} v-vpls customer 1 vlan {vlan} sap lt:1/1/3:{vlan}\n")
        time.sleep(0.5)
        shell.send("configure service vpls {vlan} v-vpls customer 1 vlan {vlan} sap lt:1/1/4:{vlan}\n")
        time.sleep(0.5)
        shell.send("configure service vpls {vlan} v-vpls customer 1 vlan {vlan} sap lt:1/1/5:{vlan}\n")
        time.sleep(0.5)
        shell.send("configure service vpls {vlan} v-vpls customer 1 vlan {vlan} sap lt:1/1/6:{vlan}\n")
        time.sleep(0.5)
        shell.send("configure service vpls {vlan} v-vpls customer 1 vlan {vlan} sap lt:1/1/7:{vlan}\n")
        time.sleep(0.5)
        shell.send("configure service vpls {vlan} v-vpls customer 1 vlan {vlan} sap lt:1/1/10:{vlan}\n")
        time.sleep(0.5)
        shell.send("configure service vpls {vlan} v-vpls customer 1 vlan {vlan} sap lt:1/1/11:{vlan}\n")
        time.sleep(0.5)
        shell.send("configure service vpls {vlan} v-vpls customer 1 vlan {vlan} sap lt:1/1/12:{vlan}\n")
        time.sleep(0.5)
        shell.send("configure service vpls {vlan} v-vpls customer 1 vlan {vlan} sap lt:1/1/13:{vlan}\n")
        time.sleep(0.5)
        shell.send("configure service vpls {vlan} v-vpls customer 1 vlan {vlan} sap lt:1/1/14:{vlan}\n")
        time.sleep(0.5)
        shell.send("configure service vpls {vlan} v-vpls customer 1 vlan {vlan} sap lt:1/1/15:{vlan}\n")
        time.sleep(0.5)
        shell.send("configure service vpls {vlan} v-vpls customer 1 vlan {vlan} sap lt:1/1/16:{vlan}\n")
        time.sleep(0.5)
        shell.send("configure service vpls {vlan} v-vpls customer 1 vlan {vlan} sap lag-1:{vlan}\n")
        time.sleep(0.5)
        shell.send("configure service vpls {vlan} v-vpls customer 1 vlan {vlan} no shutdown\n")
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
devices_df = pd.read_excel(r"D:\Earthlink\New folder\Notes\py\dataN-1.xlsx", sheet_name='Nokia')

# Create a thread pool with a maximum of 217 threads
pool = ThreadPoolExecutor(max_workers=217)

# Loop through the DataFrame rows to open sessions for each device
for _, device_row in devices_df.iterrows():
    hostname = device_row['hostname']
    username = 'absadi'
    password = 'rH6kTgyiOVEkvI4'
    port = '22'
    olt_name = device_row['olt name']
    vlan = device_row['Vlan']

    # Submit each device task to the thread pool
    pool.submit(handle_session, hostname, username, password, port, olt_name)

# Wait for all tasks to complete
pool.shutdown(wait=True)