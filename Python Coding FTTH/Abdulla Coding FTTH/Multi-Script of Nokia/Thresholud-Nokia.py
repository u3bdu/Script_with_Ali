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

        for gpon in range(1, 17):
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/1/{gpon} utilization pon-pmcollect pm-enable\n")
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/1/{gpon} utilization threshold\n")
            shell.send('txmcutilhi 100\n')
            shell.send('txmcutilmd 100\n')
            shell.send('txmcutillo 100\n')
            shell.send('txtotutilhi 100\n')
            shell.send('txtotutilmd 100\n')
            shell.send('txtotutillo 100\n')
            shell.send('rxtotutilhi 100\n')
            shell.send('rxtotutilmd 100\n')
            shell.send('rxtotutillo 100\n')
            shell.send('exit all\n')
            time.sleep(1)
        
        for gpon in range(1, 17):
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/2/{gpon} utilization pon-pmcollect pm-enable\n")
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/2/{gpon} utilization threshold\n")
            shell.send('txmcutilhi 100\n')
            shell.send('txmcutilmd 100\n')
            shell.send('txmcutillo 100\n')
            shell.send('txtotutilhi 100\n')
            shell.send('txtotutilmd 100\n')
            shell.send('txtotutillo 100\n')
            shell.send('rxtotutilhi 100\n')
            shell.send('rxtotutilmd 100\n')
            shell.send('rxtotutillo 100\n')
            shell.send('exit all\n')
            time.sleep(1)

        for gpon in range(1, 17):
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/3/{gpon} utilization pon-pmcollect pm-enable\n")
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/3/{gpon} utilization threshold\n")
            shell.send('txmcutilhi 100\n')
            shell.send('txmcutilmd 100\n')
            shell.send('txmcutillo 100\n')
            shell.send('txtotutilhi 100\n')
            shell.send('txtotutilmd 100\n')
            shell.send('txtotutillo 100\n')
            shell.send('rxtotutilhi 100\n')
            shell.send('rxtotutilmd 100\n')
            shell.send('rxtotutillo 100\n')
            shell.send('exit all\n')
            time.sleep(1)
        
        for gpon in range(1, 17):
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/4/{gpon} utilization pon-pmcollect pm-enable\n")
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/4/{gpon} utilization threshold\n")
            shell.send('txmcutilhi 100\n')
            shell.send('txmcutilmd 100\n')
            shell.send('txmcutillo 100\n')
            shell.send('txtotutilhi 100\n')
            shell.send('txtotutilmd 100\n')
            shell.send('txtotutillo 100\n')
            shell.send('rxtotutilhi 100\n')
            shell.send('rxtotutilmd 100\n')
            shell.send('rxtotutillo 100\n')
            shell.send('exit all\n')
            time.sleep(1)

        for gpon in range(1, 17):
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/5/{gpon} utilization pon-pmcollect pm-enable\n")
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/5/{gpon} utilization threshold\n")
            shell.send('txmcutilhi 100\n')
            shell.send('txmcutilmd 100\n')
            shell.send('txmcutillo 100\n')
            shell.send('txtotutilhi 100\n')
            shell.send('txtotutilmd 100\n')
            shell.send('txtotutillo 100\n')
            shell.send('rxtotutilhi 100\n')
            shell.send('rxtotutilmd 100\n')
            shell.send('rxtotutillo 100\n')
            shell.send('exit all\n')
            time.sleep(1)
        
        for gpon in range(1, 17):
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/6/{gpon} utilization pon-pmcollect pm-enable\n")
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/6/{gpon} utilization threshold\n")
            shell.send('txmcutilhi 100\n')
            shell.send('txmcutilmd 100\n')
            shell.send('txmcutillo 100\n')
            shell.send('txtotutilhi 100\n')
            shell.send('txtotutilmd 100\n')
            shell.send('txtotutillo 100\n')
            shell.send('rxtotutilhi 100\n')
            shell.send('rxtotutilmd 100\n')
            shell.send('rxtotutillo 100\n')
            shell.send('exit all\n')
            time.sleep(1)

        for gpon in range(1, 17):
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/7/{gpon} utilization pon-pmcollect pm-enable\n")
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/7/{gpon} utilization threshold\n")
            shell.send('txmcutilhi 100\n')
            shell.send('txmcutilmd 100\n')
            shell.send('txmcutillo 100\n')
            shell.send('txtotutilhi 100\n')
            shell.send('txtotutilmd 100\n')
            shell.send('txtotutillo 100\n')
            shell.send('rxtotutilhi 100\n')
            shell.send('rxtotutilmd 100\n')
            shell.send('rxtotutillo 100\n')
            shell.send('exit all\n')
            time.sleep(1)

        for gpon in range(1, 17):
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/10/{gpon} utilization pon-pmcollect pm-enable\n")
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/10/{gpon} utilization threshold\n")
            shell.send('txmcutilhi 100\n')
            shell.send('txmcutilmd 100\n')
            shell.send('txmcutillo 100\n')
            shell.send('txtotutilhi 100\n')
            shell.send('txtotutilmd 100\n')
            shell.send('txtotutillo 100\n')
            shell.send('rxtotutilhi 100\n')
            shell.send('rxtotutilmd 100\n')
            shell.send('rxtotutillo 100\n')
            shell.send('exit all\n')
            time.sleep(1)
        
        for gpon in range(1, 17):
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/11/{gpon} utilization pon-pmcollect pm-enable\n")
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/11/{gpon} utilization threshold\n")
            shell.send('txmcutilhi 100\n')
            shell.send('txmcutilmd 100\n')
            shell.send('txmcutillo 100\n')
            shell.send('txtotutilhi 100\n')
            shell.send('txtotutilmd 100\n')
            shell.send('txtotutillo 100\n')
            shell.send('rxtotutilhi 100\n')
            shell.send('rxtotutilmd 100\n')
            shell.send('rxtotutillo 100\n')
            shell.send('exit all\n')
            time.sleep(1)

        for gpon in range(1, 17):
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/12/{gpon} utilization pon-pmcollect pm-enable\n")
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/12/{gpon} utilization threshold\n")
            shell.send('txmcutilhi 100\n')
            shell.send('txmcutilmd 100\n')
            shell.send('txmcutillo 100\n')
            shell.send('txtotutilhi 100\n')
            shell.send('txtotutilmd 100\n')
            shell.send('txtotutillo 100\n')
            shell.send('rxtotutilhi 100\n')
            shell.send('rxtotutilmd 100\n')
            shell.send('rxtotutillo 100\n')
            shell.send('exit all\n')
            time.sleep(1)
        
        for gpon in range(1, 17):
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/13/{gpon} utilization pon-pmcollect pm-enable\n")
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/13/{gpon} utilization threshold\n")
            shell.send('txmcutilhi 100\n')
            shell.send('txmcutilmd 100\n')
            shell.send('txmcutillo 100\n')
            shell.send('txtotutilhi 100\n')
            shell.send('txtotutilmd 100\n')
            shell.send('txtotutillo 100\n')
            shell.send('rxtotutilhi 100\n')
            shell.send('rxtotutilmd 100\n')
            shell.send('rxtotutillo 100\n')
            shell.send('exit all\n')
            time.sleep(1)

        for gpon in range(1, 17):
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/14/{gpon} utilization pon-pmcollect pm-enable\n")
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/14/{gpon} utilization threshold\n")
            shell.send('txmcutilhi 100\n')
            shell.send('txmcutilmd 100\n')
            shell.send('txmcutillo 100\n')
            shell.send('txtotutilhi 100\n')
            shell.send('txtotutilmd 100\n')
            shell.send('txtotutillo 100\n')
            shell.send('rxtotutilhi 100\n')
            shell.send('rxtotutilmd 100\n')
            shell.send('rxtotutillo 100\n')
            shell.send('exit all\n')
            time.sleep(1)
        
        for gpon in range(1, 17):
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/15/{gpon} utilization pon-pmcollect pm-enable\n")
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/15/{gpon} utilization threshold\n")
            shell.send('txmcutilhi 100\n')
            shell.send('txmcutilmd 100\n')
            shell.send('txmcutillo 100\n')
            shell.send('txtotutilhi 100\n')
            shell.send('txtotutilmd 100\n')
            shell.send('txtotutillo 100\n')
            shell.send('rxtotutilhi 100\n')
            shell.send('rxtotutilmd 100\n')
            shell.send('rxtotutillo 100\n')
            shell.send('exit all\n')
            time.sleep(1)

        for gpon in range(1, 17):
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/16/{gpon} utilization pon-pmcollect pm-enable\n")
            time.sleep(0.5)
            shell.send(f"configure pon interface 1/1/16/{gpon} utilization threshold\n")
            shell.send('txmcutilhi 100\n')
            shell.send('txmcutilmd 100\n')
            shell.send('txmcutillo 100\n')
            shell.send('txtotutilhi 100\n')
            shell.send('txtotutilmd 100\n')
            shell.send('txtotutillo 100\n')
            shell.send('rxtotutilhi 100\n')
            shell.send('rxtotutilmd 100\n')
            shell.send('rxtotutillo 100\n')
            shell.send('exit all\n')
            time.sleep(1)
        
        time.sleep(10)
        shell.send('configure qos profiles bandwidth GPONProfile committed-info-rate 64 assured-info-rate 64 excessive-info-rate 1228544\n')
        time.sleep(0.5)
        shell.send('delay-tolerance 32\n')
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

    # Submit each device task to the thread pool
    pool.submit(handle_session, hostname, username, password, port, olt_name)

# Wait for all tasks to complete
pool.shutdown(wait=True)