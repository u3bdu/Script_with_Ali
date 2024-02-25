#لاتستخراج بورات اليوزرية لمجموعة من الاولتيات حسب المعطيات في الشيت الذي يتم القراءة منه واحد تلو الاخر او سوية حسب الاختيار

import pandas as pd
import paramiko
import time
from concurrent.futures import ThreadPoolExecutor
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Read Excel file using pandas
df = pd.read_excel(r'HIP.xlsx')

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
            time.sleep(5)




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

# Create a thread pool with a maximum of 4 threads
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
all_data.to_excel('Saad12.xlsx', index=False)