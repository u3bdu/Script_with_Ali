#تغيير فيلانات اليوزرية


import paramiko
import time
from concurrent.futures import ThreadPoolExecutor
import pandas as pd

# Define device information
device = {
    "hostname": "10.6.181.115",
    "username": "admin",
    "password": "vT1xC7uB4iG0lM8k@",
    "port": 22,
}


# Create SSH client object
ssh = paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

# Establish SSH connection to the device
ssh.connect(**device)

# Create shell object to interact with the device
shell = ssh.invoke_shell()

# Send command to device to set idle-timeout to 1000 sec
shell.send("configure system security ssh server-profile idle-timeout 1000\n")
time.sleep(5)

# Close SSH connection
ssh.close()

# Create SSH client object
ssh = paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

# Establish SSH connection to the device
ssh.connect(**device)

# Create shell object to interact with the device
shell = ssh.invoke_shell()

# Send command to device to set idle-timeout to 60 sec
shell.send("configure system security ssh server-profile idle-timeout 60\n")
time.sleep(2)
shell.send("exit all\n")

# Read the VLAN information for the specific OLT from the corresponding sheet
vlan_df = pd.read_excel("Data Storage1.xlsx", sheet_name="Data")

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
    time.sleep(0.50)
    print(port)

# Wait for the commands to be processed and the output (optional)
time.sleep(1)

# Close SSH connection
ssh.close()