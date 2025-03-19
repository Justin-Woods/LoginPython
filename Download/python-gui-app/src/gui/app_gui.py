from tkinter import Tk, Label, Entry, Button, StringVar, filedialog, messagebox
from tkinter.ttk import Progressbar
import os
from utils.downloader import download_file, extract_zip
import threading
import socket
import shutil

def get_computer_name():
    return socket.gethostname()

def get_school_drive():
    computer_name = get_computer_name()
    name_split = computer_name.split("-")
    if name_split:
        return f"\\\\ad.ccrsb.ca\\xadmin-{name_split[0]}"
    return ""
def get_current_username():
    return os.getlogin()

username = get_current_username()

DEFAULT_SCRIPT_LOCATION = f"\\\\ad.ccrsb.ca\\it-home\\IT-SCHOOL-HOME\\{username}"
DEFAULT_WORKSTATION_PATH = r"C:\CCRCE"
DEFAULT_SHARED_STORAGE = get_school_drive()

class AppGUI:
    def __init__(self, master):
        self.master = master
        master.title("Download and Extract Tool")

        self.script_location_var = StringVar(value=DEFAULT_SCRIPT_LOCATION)
        self.workstation_path_var = StringVar(value=DEFAULT_WORKSTATION_PATH)
        self.shared_storage_var = StringVar(value=DEFAULT_SHARED_STORAGE)

        Label(master, text="Script Location:").grid(row=0, column=0, padx=10, pady=10)
        self.script_location_entry = Entry(master, textvariable=self.script_location_var, width=50)
        self.script_location_entry.grid(row=0, column=1, padx=10, pady=10)
        Button(master, text="Browse", command=self.browse_script_location).grid(row=0, column=2, padx=10, pady=10)

        Label(master, text="Workstation Path:").grid(row=1, column=0, padx=10, pady=10)
        self.workstation_path_entry = Entry(master, textvariable=self.workstation_path_var, width=50)
        self.workstation_path_entry.grid(row=1, column=1, padx=10, pady=10)
        Button(master, text="Browse", command=self.browse_workstation_path).grid(row=1, column=2, padx=10, pady=10)

        Label(master, text="Shared Storage:").grid(row=2, column=0, padx=10, pady=10)
        self.shared_storage_entry = Entry(master, textvariable=self.shared_storage_var, width=50)
        self.shared_storage_entry.grid(row=2, column=1, padx=10, pady=10)
        Button(master, text="Browse", command=self.browse_shared_storage).grid(row=2, column=2, padx=10, pady=10)

        self.download_button = Button(master, text="Download", command=self.start_download)
        self.download_button.grid(row=3, column=1, pady=20)

        self.progress = Progressbar(master, orient="horizontal", length=300, mode="determinate")
        self.progress.grid(row=4, column=0, columnspan=3, padx=10, pady=10)

    def browse_script_location(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.script_location_var.set(folder_selected)

    def browse_workstation_path(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.workstation_path_var.set(folder_selected)

    def browse_shared_storage(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.shared_storage_var.set(folder_selected)

    def start_download(self):
        # Disable the download button to prevent multiple clicks
        self.download_button.config(state="disabled")
        
        # Start the download and extraction in a separate thread
        thread = threading.Thread(target=self.download_and_extract)
        thread.start()

    def download_and_extract(self):
        url = "https://github.com/Justin-Woods/LoginPython/archive/refs/heads/main.zip"
        dest_path = os.path.join(self.workstation_path_var.get(), "Login.zip")
        extract_to = self.workstation_path_var.get()
        script_location = self.script_location_var.get()

        try:
            # Download the file with progress
            self.download_file_with_progress(url, dest_path)
            
            # Extract the zip file with progress
            self.extract_zip_with_progress(dest_path)
            
            # Copy the extracted folder to the script location
            extracted_folder = os.path.join(extract_to, "LoginPython-main")
            if os.path.exists(extracted_folder):
                shutil.copytree(extracted_folder, script_location, dirs_exist_ok=True)
                shutil.rmtree(extracted_folder)  # Clean up the extracted folder

            messagebox.showinfo("Success", "Download, extraction, and copy completed successfully.")
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            # Re-enable the download button after the process is complete
            self.download_button.config(state="normal")

    def download_file_with_progress(self, url, dest_path):
        response = download_file(url, dest_path, progress_callback=self.update_progress)
        return response

    def extract_zip_with_progress(self, zip_path):
        extract_zip(zip_path, self.workstation_path_var.get(), progress_callback=self.update_progress)

    def update_progress(self, current, total):
        if total > 0:  # Avoid division by zero
            progress = (current / total) * 100  # Calculate percentage
            self.progress['value'] = progress
            self.master.update_idletasks()
        else:
            self.progress['value'] = 0  # Set progress to 0 if total is invalid
            self.master.update_idletasks()

if __name__ == "__main__":
    root = Tk()
    app = AppGUI(root)
    root.mainloop()