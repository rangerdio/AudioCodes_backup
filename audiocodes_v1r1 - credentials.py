# Beta Working script
import pandas as pd     # exel
import paramiko
import os
from datetime import datetime
from colorama import Fore, Style, init as colorama_init
# import socket
import sys
colorama_init()

error_flag = False

# getting paths
script_directory = os.path.dirname(sys.argv[0])
log_file_path = os.path.join(script_directory, "AudioCodes_backup.log")

# username = 'Admin'
# password = 'Admin'

# get target devices and credentials from Excel file
excel_file = pd.read_excel(f'{script_directory}/devices.xlsx')

device_list = excel_file.iloc[:, 0].tolist()
usernames = excel_file.iloc[:, 1].tolist()
passwords = excel_file.iloc[:, 2].tolist()


# Remove the "devices/username/password" from device list
# if device_list[0].lower() == "devices":
#     device_list[0] = "devices"
#     device_list.remove("devices")
# if usernames[0].lower() == "usernames":
#     usernames[0] = "usernames"
#     usernames.remove("usernames")
# if passwords[0].lower() == "passwords":
#     passwords[0] = "passwords"
#     passwords.remove("passwords")
def normalize_excel_data(excel_column, keyword):
    if excel_column[0].lower() == keyword:
        excel_column[0] = keyword
        excel_column.remove(keyword)


normalize_excel_data(device_list, "devices")
normalize_excel_data(usernames, "usernames")
normalize_excel_data(passwords, "passwords")


with open(log_file_path, "a") as log_file:
    def log_print(*args, **kwargs):
        # print to both console and file
        print(*args, **kwargs)
        print(*args, file=log_file, **kwargs)

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_print(f"{timestamp} - {Fore.CYAN} ------------------------------- {Style.RESET_ALL}")
    log_print(f"{timestamp} - {Fore.CYAN}           Start Backup           {Style.RESET_ALL}")
    log_print(f"{timestamp} - {Fore.CYAN} ------------------------------- {Style.RESET_ALL}")

    log_print(f"{timestamp} - Log file path: {log_file_path}")
    log_print(f"{timestamp} - Download location: {script_directory}")
    log_print(f"{timestamp} - Backup configuration from devices: {', '.join(device_list)}")

    for current_device in device_list:
        current_username = usernames[device_list.index(current_device)]
        current_password = passwords[device_list.index(current_device)]
        # --------------------------------------------------------------------------------- #
        # device list must be with unique devices due to credential index problems above #
        # if not unique, the user/password for lowest index will always be taken            #
        # this issue will be fixed later                                                    #
        #  I can use sets, to compare the list with set to confirm this, or just use sets   #
        # --------------------------------------------------------------------------------- #
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_print(f"{timestamp} -")
        log_print(f"{timestamp} - Current device is {current_device}")
        # print(current_device, current_index, usernames[current_index], passwords[current_index])
        try:
            # connect to the current device using paramiko
            ssh = paramiko.SSHClient()
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            ssh.connect(hostname=current_device, username=current_username, password=current_password)

            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_print(f"{timestamp} -  Open ssh to {current_device}")

            # define remote and local paths based on script
            remote_file_path = '/configuration/ini.ini'  # (this is ini location in AudioCodes)
            file_name_timestamp = datetime.now().strftime("%Y_%m_%d_%H%M%S")
            local_file_path = os.path.join(script_directory, f"{current_device}_{file_name_timestamp}_ini.ini")

            # open an SFTP session on the SSH connection
            sftp = ssh.open_sftp()
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_print(f"{timestamp} -    Open SFTP to {current_device}")

            # download the file
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_print(f"{timestamp} -      Downloading configuration file from {current_device}")
            with sftp.file(remote_file_path, 'rb') as remote_file:
                with open(local_file_path, 'wb') as local_file:
                    local_file.write(remote_file.read())
                    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    log_print(f"{timestamp} -      Downloaded {current_device}{remote_file_path} to {local_file_path}")

            # close the SFTP session
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_print(f"{timestamp} -    Closing SFTP connection from {current_device}")
            sftp.close()

            # close the SSH connection
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_print(f"{timestamp} -  Closing ssh connection from {current_device}")
            ssh.close()

        except paramiko.AuthenticationException as e:
            error_flag = True
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_print(f"{timestamp} - {Fore.CYAN}Authentication error for {current_device}: {e}{Style.RESET_ALL}")
        except paramiko.SSHException as e:
            error_flag = True
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_print(f"{timestamp} - {Fore.CYAN}SSH error for {current_device}: {e}{Style.RESET_ALL}")
        except paramiko.SFTPError as e:
            error_flag = True
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_print(f"{timestamp} - {Fore.CYAN}SFTP error for {current_device}: {e}{Style.RESET_ALL}")
        except Exception as e:
            error_flag = True
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_print(f"{timestamp} - {Fore.CYAN}Unexpected error for {current_device}: {e}{Style.RESET_ALL}")

    log_print(f"{timestamp} - {Fore.CYAN} ---------------------------------- {Style.RESET_ALL}")
    if error_flag:
        log_print(f"{timestamp} - {Fore.CYAN}      Backup Done with errors           {Style.RESET_ALL}")
    else:
        log_print(f"{timestamp} - {Fore.CYAN}             Backup Done           {Style.RESET_ALL}")
    log_print(f"{timestamp} - {Fore.CYAN} ---------------------------------- {Style.RESET_ALL}")
    log_file.close()
""" 
exit_ = input(f'Press {Fore.CYAN}"Enter"{Style.RESET_ALL} to close the app: ')
"""
