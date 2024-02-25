import pandas as pd
import paramiko
import time
from concurrent.futures import ThreadPoolExecutor
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook

# Read Excel file using pandas
df = pd.read_excel(r'D:\Earthlink\New folder\Notes\py\Rings\Data and script\Windows\Excel\RBG27.xlsx', sheet_name="FDTs")

def handle_session(hostname, username, password, port,Ring, FDT_Name, MG1, MG2, MG3, MG4, MG5, MG6, MG7, MG8, MG9, MG10, MG11, MG12, BG1, BG2, BG3, BG4, BG5, BG6, BG7, BG8, BG9, BG10, BG11, BG12,seq):
    max_retries = 3  # Maximum number of SSH connection retries
    retry_delay = 5  # Delay between SSH connection retries (in seconds)

    for attempt in range(max_retries):
        try:

            print(Ring)

            # Create SSH client object
            ssh = paramiko.SSHClient()
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            # Establish SSH connection
            ssh.connect(hostname=hostname, username=username, password=password, port=port)

            # Create an empty DataFrame to store the dataa
            all_data = pd.DataFrame()



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

            shell.send(f"{FDT_Name+'B'}\n")
            time.sleep(0.50)

            shell.send("configure system security ssh server-profile idle-timeout 60\n")
            time.sleep(1)

            print(FDT_Name)

            # Backup Up
            shell.send(f"configure pon interface 1/1/{BG1} admin-state up\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{BG2} admin-state up\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{BG3} admin-state up\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{BG4} admin-state up\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{BG5} admin-state up\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{BG6} admin-state up\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{BG7} admin-state up\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{BG8} admin-state up\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{BG9} admin-state up\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{BG10} admin-state up\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{BG11} admin-state up\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{BG12} admin-state up\n")
            time.sleep(0.25)

            # Main down
            shell.send(f"configure pon interface 1/1/{MG1} admin-state down\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{MG2} admin-state down\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{MG3} admin-state down\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{MG4} admin-state down\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{MG5} admin-state down\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{MG6} admin-state down\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{MG7} admin-state down\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{MG8} admin-state down\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{MG9} admin-state down\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{MG10} admin-state down\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{MG11} admin-state down\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{MG12} admin-state down\n")
            time.sleep(0.25)
            shell.send("exit all\n")

            time.sleep(30)

            # Power Optics Backup
            for port in range(1, 16):
                time.sleep(0.5)
                shell.send(f"show equipment ont optics 1/1/{MG1}/{port}\n")
            for port in range(1, 16):
                time.sleep(0.5)
                shell.send(f"show equipment ont optics 1/1/{MG2}/{port}\n")
            for port in range(1, 16):
                time.sleep(0.5)
                shell.send(f"show equipment ont optics 1/1/{MG3}/{port}\n")
            for port in range(1, 16):
                time.sleep(0.5)
                shell.send(f"show equipment ont optics 1/1/{MG4}/{port}\n")
            for port in range(1, 16):
                time.sleep(0.5)
                shell.send(f"show equipment ont optics 1/1/{MG5}/{port}\n")
            for port in range(1, 16):
                time.sleep(0.5)
                shell.send(f"show equipment ont optics 1/1/{MG6}/{port}\n")
            for port in range(1, 16):
                time.sleep(0.5)
                shell.send(f"show equipment ont optics 1/1/{MG7}/{port}\n")     
            for port in range(1, 16):
                time.sleep(0.5)
                shell.send(f"show equipment ont optics 1/1/{MG8}/{port}\n")
            for port in range(1, 16):
                time.sleep(2)
                shell.send(f"show equipment ont optics 1/1/{MG9}/{port}\n")
            for port in range(1, 16):
                time.sleep(2)
                shell.send(f"show equipment ont optics 1/1/{MG10}/{port}\n")                
            for port in range(1, 16):
                time.sleep(2)
                shell.send(f"show equipment ont optics 1/1/{MG11}/{port}\n")                
            for port in range(1, 16):
                time.sleep(2)
                shell.send(f"show equipment ont optics 1/1/{MG12}/{port}\n")

            # Main Up
            shell.send(f"configure pon interface 1/1/{MG1} admin-state up\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{MG2} admin-state up\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{MG3} admin-state up\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{MG4} admin-state up\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{MG5} admin-state up\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{MG6} admin-state up\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{MG7} admin-state up\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{MG8} admin-state up\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{MG9} admin-state up\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{MG10} admin-state up\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{MG11} admin-state up\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{MG12} admin-state up\n")
            time.sleep(0.25)
            shell.send("exit all\n")
            time.sleep(0.25)

            time.sleep(5)

            # Receive the output for display current-configuration
            output1 = b""
            while True:
                if shell.recv_ready():
                    buffer = shell.recv(1024)
                    output1 += buffer
                else:
                    break

            # Decode using 'latin-1' encoding
            output1 = output1.decode('latin-1')

            # Split the output into lines
            lines1 = output1.splitlines()

            # Parse the output for Sheet1
            data_output1 = [line.split() for line in lines1 if line.strip()]
            time.sleep(2)

            shell.send(f" {FDT_Name+'M'}")
            time.sleep(0.50)
            
            # Backup Down
            shell.send(f"configure pon interface 1/1/{BG1} admin-state down\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{BG2} admin-state down\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{BG3} admin-state down\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{BG4} admin-state down\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{BG5} admin-state down\n")
            time.sleep(0.25)
            shell.send(f"configure pon interface 1/1/{BG6} admin-state down\n")
            time.sleep(1)
            shell.send(f"configure pon interface 1/1/{BG7} admin-state down\n")
            time.sleep(1)
            shell.send(f"configure pon interface 1/1/{BG8} admin-state down\n")
            time.sleep(1)
            shell.send(f"configure pon interface 1/1/{BG9} admin-state down\n")
            time.sleep(1)
            shell.send(f"configure pon interface 1/1/{BG10} admin-state down\n")
            time.sleep(1)
            shell.send(f"configure pon interface 1/1/{BG11} admin-state down\n")
            time.sleep(1)
            shell.send(f"configure pon interface 1/1/{BG12} admin-state down\n")
            time.sleep(1)
            shell.send("exit all\n")

            time.sleep(30)

            # Power Optics Main
            for port in range(1, 16):
                time.sleep(0.5)
                shell.send(f"show equipment ont optics 1/1/{MG1}/{port}\n")
            for port in range(1, 16):
                time.sleep(0.5)
                shell.send(f"show equipment ont optics 1/1/{MG2}/{port}\n")
            for port in range(1, 16):
                time.sleep(0.5)
                shell.send(f"show equipment ont optics 1/1/{MG3}/{port}\n")
            for port in range(1, 16):
                time.sleep(0.5)
                shell.send(f"show equipment ont optics 1/1/{MG4}/{port}\n")
            for port in range(1, 16):
                time.sleep(0.5)
                shell.send(f"show equipment ont optics 1/1/{MG5}/{port}\n")
            for port in range(1, 16):
                time.sleep(0.5)
                shell.send(f"show equipment ont optics 1/1/{MG6}/{port}\n")
            for port in range(1, 16):
                time.sleep(0.5)
                shell.send(f"show equipment ont optics 1/1/{MG7}/{port}\n")     
            for port in range(1, 16):
                time.sleep(0.5)
                shell.send(f"show equipment ont optics 1/1/{MG8}/{port}\n")
            for port in range(1, 16):
                time.sleep(2)
                shell.send(f"show equipment ont optics 1/1/{MG9}/{port}\n")
            for port in range(1, 16):
                time.sleep(2)
                shell.send(f"show equipment ont optics 1/1/{MG10}/{port}\n")                
            for port in range(1, 16):
                time.sleep(2)
                shell.send(f"show equipment ont optics 1/1/{MG11}/{port}\n")                
            for port in range(1, 16):
                time.sleep(2)
                shell.send(f"show equipment ont optics 1/1/{MG12}/{port}\n")

            # Backup UP
            shell.send(f"configure pon interface 1/1/{BG1} admin-state up\n")
            time.sleep(0.25)
            shell.send("exit all\n")
            time.sleep(1)
            shell.send(f"configure pon interface 1/1/{BG2} admin-state up\n")
            time.sleep(0.25)
            shell.send("exit all\n")
            time.sleep(1)
            shell.send(f"configure pon interface 1/1/{BG3} admin-state up\n")
            time.sleep(0.25)
            shell.send("exit all\n")
            time.sleep(1)
            shell.send(f"configure pon interface 1/1/{BG4} admin-state up\n")
            time.sleep(0.25)
            shell.send("exit all\n")
            time.sleep(1)
            shell.send(f"configure pon interface 1/1/{BG5} admin-state up\n")
            time.sleep(0.25)
            shell.send("exit all\n")
            time.sleep(1)
            shell.send(f"configure pon interface 1/1/{BG6} admin-state up\n")
            time.sleep(0.25)
            shell.send("exit all\n")
            time.sleep(1)
            shell.send(f"configure pon interface 1/1/{BG7} admin-state up\n")
            time.sleep(0.25)
            shell.send("exit all\n")
            time.sleep(1)
            shell.send(f"configure pon interface 1/1/{BG8} admin-state up\n")
            time.sleep(0.25)
            shell.send("exit all\n")
            time.sleep(1)
            shell.send(f"configure pon interface 1/1/{BG9} admin-state up\n")
            time.sleep(0.25)
            shell.send("exit all\n")
            time.sleep(1)
            shell.send(f"configure pon interface 1/1/{BG10} admin-state up\n")
            time.sleep(0.25)
            shell.send("exit all\n")
            time.sleep(1)
            shell.send(f"configure pon interface 1/1/{BG11} admin-state up\n")
            time.sleep(0.25)
            shell.send("exit all\n")
            time.sleep(1)
            shell.send(f"configure pon interface 1/1/{BG12} admin-state up\n")
            time.sleep(0.25)

            shell.send("exit all\n")
            time.sleep(2)


            # Receive the output for display protect-group
            
            output2 = b""
            while True:
                if shell.recv_ready():
                    buffer = shell.recv(1024)
                    output2 += buffer
                else:
                    break
            # Decode using 'latin-1' encoding
            output2 = output2.decode('latin-1')

            # Split the output into lines
            lines2 = output2.splitlines()

            # Parse the output for Sheet2
            data_output2 = [line.split() for line in lines2 if line.strip()]
            time.sleep(2)


            # Create a DataFrame for Sheet1
            df_output1 = pd.DataFrame(data_output1)
            df_output1 = df_output1.rename(columns={0: FDT_Name+" B"})  # Set the first column header to FDT

            # Create a DataFrame for Sheet2
            df_output2 = pd.DataFrame(data_output2)
            df_output2 = df_output2.rename(columns={0: FDT_Name+" M"})  # Set the first column header to FDT
            
            # Save the DataFrames for Sheet1 and Sheet2 to Excel
            df_output1.to_excel(writer, sheet_name='B'+str (seq), index=False)
            df_output2.to_excel(writer, sheet_name='M'+str (seq), index=False)


            #Save the data for the current device in a separate sheet
            all_data.to_excel(writer, index=False)

            #ssh close
            ssh.close()
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

# Create a thread pool with a maximum of 5 threads
pool = ThreadPoolExecutor(max_workers=5)
 
time.sleep(2) 
for _, row in df.iterrows():
    hostname = row['hostname']
    username = 'absadi'
    password = 'rH6kTgyiOVEkvI4'
    port = '22'
    Ring = row ['Ring']
    FDT_Name = row["FDT name"]
    seq = row ['seq']
    MG1 = row ['MG1']
    MG2 = row ['MG2']
    MG3 = row ['MG3']
    MG4 = row ['MG4']
    MG5 = row ['MG5']
    MG6 = row ['MG6']
    MG7 = row ['MG7']
    MG8 = row ['MG8']
    MG9 = row ['MG9']
    MG10 = row ['MG10']
    MG11 = row ['MG11']
    MG12 = row ['MG12']
    BG1 = row ['BG1']
    BG2 = row ['BG2']
    BG3 = row ['BG3']
    BG4 = row ['BG4']
    BG5 = row ['BG5']
    BG6 = row ['BG6']
    BG7 = row ['BG7']
    BG8 = row ['BG8']
    BG9 = row ['BG9']
    BG10 = row ['BG10']
    BG11 = row ['BG11']
    BG12 = row ['BG12']

    # Create an Excel writer object for each iteration
    excel_filename = rf'D:\Earthlink\New folder\Notes\py\Rings\Data and script\Windows\result\{Ring}.xlsx'
    writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')

    pool.submit(handle_session, hostname, username, password, port,Ring, FDT_Name, MG1, MG2, MG3, MG4, MG5, MG6, MG7, MG8, MG9, MG10, MG11, MG12, BG1, BG2, BG3, BG4, BG5, BG6, BG7, BG8, BG9, BG10, BG11, BG12,seq)

# Shutdown the thread pool to wait for all tasks to complete
pool.shutdown(wait=True)


# Save the Excel file
writer._save()