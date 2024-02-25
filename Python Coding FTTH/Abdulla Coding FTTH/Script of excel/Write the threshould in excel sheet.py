import time
import openpyxl
import subprocess

# Function to execute commands using subprocess
def execute_command(command):
    result = subprocess.run(command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    return result.stdout.decode(), result.stderr.decode()

# Create a new Excel workbook and select the active sheet
workbook = openpyxl.Workbook()
sheet = workbook.active

# Loop through the range and add each command to the Excel sheet
for interface in range(1, 17):
    for gpon in range(1, 17):
        command1 = f"configure pon interface 1/1/{interface}/{gpon} utilization pon-pmcollect pm-enable"
        command2 = f"configure pon interface 1/1/{interface}/{gpon} utilization threshold"
        commands = ['txmcutilhi 100', 'txmcutilmd 100', 'txmcutillo 100',
                    'txtotutilhi 100', 'txtotutilmd 100', 'txtotutillo 100',
                    'rxtotutilhi 100', 'rxtotutilmd 100', 'rxtotutillo 100', 'exit all']

        # Execute and append each command to the Excel sheet
        for cmd in [command1, command2] + commands:
            print(cmd)
            execute_command(cmd)
            
            sheet.append([cmd])  # Each command in a new row

# Save the Excel file
workbook.save("output1.xlsx")
