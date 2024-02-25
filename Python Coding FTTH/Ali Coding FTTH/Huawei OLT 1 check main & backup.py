#لتجييك بورات اليوزرية مين وباك اب لاولتي واحد

import paramiko
import time
import re
import pandas as pd

# define device information
device = {
"hostname": "10.6.181.131",
"username": "irqnbn",
"password": "VSl3DbqWVDfOd0w@",
"port": 22,
}

# create SSH client object
ssh = paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

# establish SSH connection to device
ssh.connect(**device)

# create shell object to interact with device
shell = ssh.invoke_shell()

# send command to device
shell.send("enable\n")

# send command to device
shell.send("config\n")

# wait for command output to be generated
time.sleep(2)

# send command to device to force-switch to Backup
for i in range(1, 17):
    shell.send(f"protect-group {i}\n")
    shell.send("\n")
    shell.send(f"force-switch port 0/1/{i-1} to 0/18/{i-1} \n")
    shell.send("\n")
    shell.send("quit\n")
for i in range(17, 33):
    protect_group = f"protect-group {i}\n"
    force_switch = f"force-switch port 0/2/{i-17} to 0/17/{i-17} \n"
    shell.send(protect_group)
    shell.send("\n")
    shell.send(force_switch)
    shell.send("\n")
    shell.send("quit\n")
for i in range(33, 49):
    shell.send(f"protect-group {i}\n")
    shell.send("\n")
    shell.send(f"force-switch port 0/3/{i-33} to 0/16/{i-33} \n")
    shell.send("\n")
    shell.send("quit\n")
for i in range(49, 65):
    shell.send(f"protect-group {i}\n")
    shell.send("\n")
    shell.send(f"force-switch port 0/4/{i-49} to 0/15/{i-49} \n")
    shell.send("\n")
    shell.send("quit\n")
for i in range(65, 81):
    shell.send(f"protect-group {i}\n")
    shell.send("\n")
    shell.send(f"force-switch port 0/5/{i-65} to 0/14/{i-65} \n")
    shell.send("\n")
    shell.send("quit\n")
for i in range(81, 97):
    shell.send(f"protect-group {i}\n")
    shell.send("\n")
    shell.send(f"force-switch port 0/6/{i-81} to 0/13/{i-81} \n")
    shell.send("\n")
    shell.send("quit\n")
for i in range(97, 113):
    shell.send(f"protect-group {i} \n")
    shell.send("\n")
    shell.send(f"force-switch port 0/7/{i-97} to 0/12/{i-97} \n")
    shell.send("\n")
    shell.send("quit\n")
# wait for command output to be generated
time.sleep(30)

shell.send("display ont info summary 0 | no-more\n")
shell.send("\n")
time.sleep(90)

# send command to device to force-switch to Main
for i in range(1, 17):
    shell.send(f"protect-group {i}\n")
    shell.send("\n")
    shell.send(f"force-switch port 0/18/{i-1} to 0/1/{i-1}\n")
    shell.send("\n")
    shell.send("quit\n")
for i in range(17, 33):
    shell.send(f"protect-group {i}\n")
    shell.send("\n")
    shell.send(f"force-switch port 0/17/{i-17} to 0/2/{i-17}\n")
    shell.send("\n")
    shell.send("quit\n")
for i in range(33, 49):
    shell.send(f"protect-group {i}\n")
    shell.send("\n")
    shell.send(f"force-switch port 0/16/{i-33} to 0/3/{i-33}\n")
    shell.send("\n")
    shell.send("quit\n")
for i in range(49, 65):
    shell.send(f"protect-group {i}\n")
    shell.send("\n")
    shell.send(f"force-switch port 0/15/{i-49} to 0/4/{i-49}\n")
    shell.send("\n")
    shell.send("quit\n") 
for i in range(65, 81):
    shell.send(f"protect-group {i}\n")
    shell.send("\n")
    shell.send(f"force-switch port 0/14/{i-65} to 0/5/{i-65}\n")
    shell.send("\n")
    shell.send("quit\n")
for i in range(81, 97):
    shell.send(f"protect-group {i}\n")
    shell.send("\n")
    shell.send(f"force-switch port 0/13/{i-81} to 0/6/{i-81}\n")
    shell.send("\n")
    shell.send("quit\n")
for i in range(97, 113):
    shell.send(f"protect-group {i}\n")
    shell.send("\n")
    shell.send(f"force-switch port 0/12/{i-97} to 0/7/{i-97}\n")
    shell.send("\n")
    shell.send("quit\n")

# wait for command output to be generated
time.sleep(30)

shell.send("display ont info summary 0 | no-more\n")
shell.send("\n")
time.sleep(90)

shell.send("\n")
time.sleep(2)

# send command to device to undo force-switch
for i in range(1, 17):
    shell.send(f"protect-group {i} \n")  
    shell.send("\n")
    shell.send(f"undo force-switch port 0/18/{i-1} to 0/1/{i-1}\n")
    shell.send("\n")
    shell.send("quit\n")
for i in range(17, 33):
    shell.send(f"protect-group {i} \n")
    shell.send("\n")
    shell.send(f"undo force-switch port 0/17/{i-17} to 0/2/{i-17}\n")
    shell.send("\n")
    shell.send("quit\n")
for i in range(33, 49):
    shell.send(f"protect-group {i} \n")  
    shell.send("\n")
    shell.send(f"undo force-switch port 0/16/{i-33} to 0/3/{i-33}\n")
    shell.send("\n") 
    shell.send("quit\n")
for i in range(49, 65):
    shell.send(f"protect-group {i} \n")  
    shell.send("\n")
    shell.send(f"undo force-switch port 0/15/{i-49} to 0/4/{i-49}\n")
    shell.send("\n") 
    shell.send("quit\n")
for i in range(65, 81):
    shell.send(f"protect-group {i} \n")  
    shell.send("\n")
    shell.send(f"undo force-switch port 0/14/{i-65} to 0/5/{i-65}\n")
    shell.send("\n") 
    shell.send("quit\n")
for i in range(81, 97):
    shell.send(f"protect-group {i} \n")  
    shell.send("\n")
    shell.send(f"undo force-switch port 0/13/{i-81} to 0/6/{i-81}\n")
    shell.send("\n") 
    shell.send("quit\n")
for i in range(97, 113):
    shell.send(f"protect-group {i} \n")  
    shell.send("\n")
    shell.send(f"undo force-switch port 0/12/{i-97} to 0/7/{i-97}\n")
    shell.send("\n") 
    shell.send("quit\n")

# wait for command output to be generated
time.sleep(40)

shell.send("display protect-group | no-more\n")
shell.send("\n")
time.sleep(20)

output= b""
while True:
    # Check if there is data available to be received
    if shell.recv_ready():
        # Receive the data
        buffer = shell.recv(1024)
        output += buffer
    else:
        # If no more data is available, break the loop
        break

# Decode the output
output = output.decode()
output = re.sub(r'\x1b\[1D', '', output)
print(output.encode('utf-8'))

# Parse the output and save it to the Excel sheet
lines = output.splitlines()
data = [line.split() for line in lines]

# Create a DataFrame from the data
df = pd.DataFrame(data)

# Write the DataFrame to an Excel file
df.to_excel('Nasriya2.xlsx', index=False)

# close SSH connection
ssh.close()