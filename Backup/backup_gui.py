import os
import sys
import ctypes
import shutil
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox
import logging
import tempfile

# Set up logging
temp_dir = tempfile.gettempdir()
log_file = os.path.join(temp_dir, 'backup_script.log')
logging.basicConfig(filename=log_file, level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# Function to check if the script is running with admin privileges
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

# Re-run the script with admin privileges if not already running as admin
if not is_admin():
    logging.info("Re-running script with admin privileges")
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
    sys.exit()

# Function to install pip if it's missing
def ensure_pip():
    try:
        subprocess.run([sys.executable, "-m", "ensurepip", "--default-pip"], check=True)
        logging.info("pip installed successfully")
    except subprocess.CalledProcessError as e:
        logging.error(f"Failed to install pip: {e.stderr.decode()}")
        messagebox.showerror("Error", f"Failed to install pip: {e.stderr.decode()}")

# Function to install required packages
def install(package):
    try:
        result = subprocess.run([sys.executable, "-m", "pip", "install", package], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        logging.info(f"Successfully installed {package}: {result.stdout.decode()}")
    except subprocess.CalledProcessError as e:
        logging.error(f"Failed to install {package}: {e.stderr.decode()}")
        messagebox.showerror("Error", f"Failed to install {package}: {e.stderr.decode()}")

# Ensure pip is installed
ensure_pip()

# Ensure required packages are installed
try:
    import win32api
    import win32file
except ImportError:
    logging.info("pywin32 is not installed. Installing...")
    install("pywin32")
    import win32api
    import win32file

try:
    import psutil
except ImportError:
    logging.info("psutil is not installed. Installing...")
    install("psutil")
    import psutil

# Function to check if a drive is a network drive
def is_network_drive(drive_letter):
    try:
        output = subprocess.check_output(f'net use {drive_letter} 2>NUL | FINDSTR /I "\\\\"', shell=True).decode().strip()
        return bool(output)
    except subprocess.CalledProcessError:
        return False

# Function to run the backup
def run_backup():
    backup_dir = backup_dir_var.get()
    if not backup_dir:
        messagebox.showerror("Error", "Please select a backup directory")
        return

    if not os.path.exists(backup_dir):
        os.makedirs(backup_dir)

    logging.info(f"Running backup to directory: {backup_dir}")
    backup_system_info(backup_dir)

# Function to select the backup directory
def select_backup_dir():
    default_dir = backup_dir_var.get()
    backup_dir = filedialog.askdirectory(initialdir=default_dir, title="Select Backup Directory")
    if backup_dir:
        computer_name = os.environ['COMPUTERNAME']
        backup_dir_with_sysname = os.path.join(backup_dir, computer_name)
        backup_dir_var.set(backup_dir_with_sysname)

# Function to eject the optical drive in Windows
def eject_drive(drive_letter):
    handle = win32file.CreateFile(
        f"\\\\.\\{drive_letter}:",
        win32file.GENERIC_READ,
        win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE,
        None,
        win32file.OPEN_EXISTING,
        0,
        None
    )
    FSCTL_LOCK_VOLUME = 0x90018
    FSCTL_DISMOUNT_VOLUME = 0x90020
    IOCTL_STORAGE_EJECT_MEDIA = 0x2D4808
    win32file.DeviceIoControl(handle, FSCTL_LOCK_VOLUME, None, 0, None)
    win32file.DeviceIoControl(handle, FSCTL_DISMOUNT_VOLUME, None, 0, None)
    win32file.DeviceIoControl(handle, IOCTL_STORAGE_EJECT_MEDIA, None, 0, None)
    handle.close()

# Function to check for optical drives and eject if a disk is inserted
def check_and_eject_optical_drive():
    partitions = psutil.disk_partitions(all=True)
    for partition in partitions:
        if 'cdrom' in partition.opts:
            drive_letter = partition.device.split(':')[0]
            try:
                usage = psutil.disk_usage(partition.mountpoint)
                if usage.percent > 0:
                    logging.info(f"Ejecting disk from {drive_letter}:")
                    eject_drive(drive_letter)
                else:
                    logging.info(f"No disk in {drive_letter}:")
            except Exception as e:
                logging.error(f"Error checking {drive_letter}: {e}")
                
def is_physical_drive(drive_letter):
    try:
        drive_type = subprocess.check_output(f'fsutil fsinfo drivetype {drive_letter}', shell=True).decode().strip()
        return drive_type == "Fixed Drive"
    except subprocess.CalledProcessError:
        return False
                
def backup_system_info(backup_dir=None):
    # Shows name and description
    computer_name = os.environ['COMPUTERNAME']
    print("---------------------")
    print(f"Backup {computer_name}")
    print("---------------------")
    print()

    # Converts a mapped path to a UNC path
    current_path = os.path.abspath(__file__)
    drive_letter = os.path.splitdrive(current_path)[0]
    if is_network_drive(drive_letter):
        try:
            unc_path = subprocess.check_output(f'net use {drive_letter} 2>NUL | FINDSTR /I "\\\\"', shell=True).decode().split()[-1]
            script_path = current_path.replace(drive_letter, unc_path)
        except subprocess.CalledProcessError as e:
            logging.error(f"Failed to convert to UNC path: {e.stderr.decode()}")
            script_path = current_path
    else:
        script_path = current_path

    # Backup directory
    if not backup_dir:
        print("Error: Backup directory not specified. Using current directory.")
        backup_dir = os.path.join(os.path.dirname(__file__), computer_name)
        if not os.path.exists(backup_dir):
            os.makedirs(backup_dir)
        if not os.path.exists(backup_dir):
            print(f"{os.environ['USERNAME']} does not have write access to {os.path.dirname(__file__)}.")

    # System specs
    with open(os.path.join(backup_dir, 'system.txt'), 'w') as f:
        f.write(subprocess.check_output('wmic computersystem get name,model /format:list | find "="', shell=True).decode())
        f.write(subprocess.check_output('wmic bios get serialnumber /format:list | find "="', shell=True).decode())

    # Data igloo junctions
    if os.environ['PROCESSOR_ARCHITECTURE'] == 'x86':
        subprocess.call(f'REG EXPORT "HKEY_LOCAL_MACHINE\\SOFTWARE\\Faronics\\Data Igloo\\Junctions" "{os.path.join(backup_dir, "igloo.txt")}" >NUL 2>NUL', shell=True)
    elif os.environ['PROCESSOR_ARCHITECTURE'] == 'AMD64':
        subprocess.call(f'REG EXPORT "HKEY_LOCAL_MACHINE\\SOFTWARE\\Wow6432Node\\Faronics\\Data Igloo\\Junctions" "{os.path.join(backup_dir, "igloo.txt")}" >NUL 2>NUL', shell=True)

    # OneNote documents
    if os.path.exists("C:\\OneNoteNotebooks"):
        shutil.copytree("C:\\OneNoteNotebooks", os.path.join(backup_dir, "OneNoteNotebooks"), dirs_exist_ok=True)

    # Printer information
    print("Documenting printers...")
    with open(os.path.join(backup_dir, 'printers.txt'), 'w') as f:
        f.write(subprocess.check_output('wmic printer get comment,location,name,portname | findstr /i "Name ."', shell=True).decode())

    # D drive
    if os.path.exists("D:") and is_physical_drive("D:"):
        print()
        print("D drive:")
        subprocess.call(f'ROBOCOPY "D:" "{os.path.join(backup_dir, "D")}" /e /log+:"{os.path.join(backup_dir, "backup.log")}" /fp /mt /nc /njh /njs /r:0 /tee /w:0 /xj /xf "backup.bat" "desktop.ini" "Thumbs.db" /xd "Audit" "{computer_name}" "$RECYCLE.BIN" "System Volume Information"', shell=True)
        print("D drive saved.")

    for user in os.listdir("C:\\Users"):
        skip_users = ('public', 'all users', 'default user', 'desktop.ini', 'default')
        if user.lower() in skip_users:
            continue
        print(f"{user}:")
        user_path = os.path.join("C:\\Users", user)
        if os.path.exists(os.path.join(user_path, "Desktop")):
            subprocess.call(f'ROBOCOPY "{os.path.join(user_path, "Desktop")}" "{os.path.join(backup_dir, "C", "Users", user, "Desktop")}" /e /log+:"{os.path.join(backup_dir, "backup.log")}" /fp /mt /nc /njh /njs /r:0 /tee /w:0 /xj /xd "{computer_name}" /xf "backup.bat" "desktop.ini" "Thumbs.db"', shell=True)
        chrome_bookmarks = os.path.join(user_path, 'AppData', 'Local', 'Google', 'Chrome', 'User Data', 'Default', 'Bookmarks')
        chrome_bookmarks_dest = os.path.join(backup_dir, 'C', 'Users', user, 'Bookmarks', 'Chrome')
        os.makedirs(os.path.dirname(chrome_bookmarks_dest), exist_ok=True)
        if os.path.exists(chrome_bookmarks):
            shutil.copy(chrome_bookmarks, os.path.join(backup_dir, 'C', 'Users', user, 'Bookmarks', 'Chrome'))
            
        edge_bookmarks = os.path.join(user_path, 'AppData', 'Local', 'Packages', 'Microsoft.MicrosoftEdge_8wekyb3d8bbwe', 'AC', 'MicrosoftEdge', 'User', 'Default', 'DataStore', 'Data', 'nouser1', '120712-0049', 'DBStore')
        if os.path.exists(edge_bookmarks):
            shutil.copytree(edge_bookmarks, os.path.join(backup_dir, 'C', 'Users', user, 'Bookmarks', 'Edge'), dirs_exist_ok=True)

        firefox_profiles_path = os.path.join(user_path, 'AppData', 'Roaming', 'Mozilla', 'Firefox', 'Profiles')
        if os.path.exists(firefox_profiles_path):
            firefox_profiles = [f for f in os.listdir(firefox_profiles_path)]
            for profile in firefox_profiles:
                places_sqlite = os.path.join(firefox_profiles_path, profile, 'places.sqlite')
                if os.path.exists(places_sqlite):
                    shutil.copy(places_sqlite, os.path.join(backup_dir, 'C', 'Users', user, 'Bookmarks', 'Firefox'))
        print(f"{user} saved.")
    print()
    print("Backup complete. :)")
    
# Call the function to check and eject optical drives if needed
check_and_eject_optical_drive()

# Create the main window
root = tk.Tk()
root.title("Backup Script GUI")

# Create variables
backup_dir_var = tk.StringVar(root)

# Create widgets
label = tk.Label(root, text="Backup Directory:")
label.grid(row=0, column=0, padx=5, pady=5)

backup_dir_entry = tk.Entry(root, textvariable=backup_dir_var, width=50)
backup_dir_entry.grid(row=0, column=1, padx=5, pady=5)

browse_button = tk.Button(root, text="Browse", command=select_backup_dir)
browse_button.grid(row=0, column=2, padx=5, pady=5)

backup_button = tk.Button(root, text="Run Backup", command=run_backup)
backup_button.grid(row=1, column=1, padx=5, pady=5)

# Wrap the main loop in a try-except block to catch any exceptions
try:
    logging.info("Starting GUI")
    root.mainloop()
except Exception as e:
    logging.error(f"An error occurred: {e}")
    messagebox.showerror("Error", f"An error occurred: {e}")