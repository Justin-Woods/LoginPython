import sqlite3
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import socket
import configparser
import subprocess
import ctypes
import shutil
import sys
import time
import webbrowser
sys.stdout.write("Loading please wait...\n")

# Directorys
script_directory = os.path.dirname(os.path.abspath(__file__))
db_directory = "X:/Audit/Databases"
ini_directory = os.path.join(script_directory, "settings.ini")
audit_directory = os.path.join(script_directory, "Audit/Audit.ps1")
export_script_path = os.path.join(script_directory, "Audit/Audit-Export.ps1")
admin_audit_script_path = os.path.join(script_directory, "Audit/Admin-Audit.ps1")
logs_directory = "C:/CCRCE/Logs"

# Read settings from ini file
config = configparser.ConfigParser()
config.read(ini_directory)
db_path = ""

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

if db_directory.startswith("X:"):
    db_directory = db_directory.replace("X:", get_school_drive(), 1)
db_directory = db_directory if os.path.exists(db_directory) else os.path.join(script_directory, "Audit/Databases")
#sys.stdout.write("db_directory: " + db_directory + "\n")

if admin_audit_script_path.startswith("H:"):
    username = get_current_username()
    admin_audit_script_path = admin_audit_script_path.replace("H:", f"\\\\ad.ccrsb.ca\\it-home\\IT-SCHOOL-HOME\\{username}", 1)

def fetch_databases():
    """Fetch all .db files in the specified directory."""
    return [f for f in os.listdir(db_directory) if f.endswith(".db")]

def fetch_data():
    """Fetch data from the 'SystemInfo' table."""
    if not db_path:
        return [], []
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM SystemInfo")
        data = cursor.fetchall()
        columns = [desc[0] for desc in cursor.description]
        conn.close()
        return columns, data
    except sqlite3.Error as e:
        messagebox.showerror("Database Error", str(e))
        return [], []
        
def fetch_admin_data():
    """Fetch data from the 'AdminSystemInfo' table."""
    if not db_path:
        return [], []
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM AdminSystemInfo")
        data = cursor.fetchall()
        columns = [desc[0] for desc in cursor.description]
        conn.close()
        return columns, data
    except sqlite3.Error as e:
        messagebox.showerror("Database Error", db_path + str(e))
        return [], []

def check_admin_table_exists():
    """Check if the 'AdminSystemInfo' table exists in the database."""
    if not db_path:
        return False
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='AdminSystemInfo'")
        table_exists = cursor.fetchone() is not None
        conn.close()
        return table_exists
    except sqlite3.Error as e:
        messagebox.showerror("Database Error", str(e))
        return False
    
def check_SystemInfo_exists():
    """Check if the 'SystemInfo' table exists in the database."""
    if not db_path:
        return False
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='SystemInfo'")
        table_exists = cursor.fetchone() is not None
        conn.close()
        if not table_exists and os.path.exists(db_path):
            sys.stdout.write("Database issue detected. Deleting database file...\n" + db_path + "\n")
            os.remove(db_path)
            refresh_databases()
        return table_exists
    except sqlite3.Error as e:
        messagebox.showerror("Database Error", str(e))
        return False
       
def display_table():
    """Display 'SystemInfo' table data with columns as rows, alternating row colors, and adjustable second column width."""
    columns, data = fetch_data()
    for widget in tree_frame.winfo_children():
        widget.destroy()
    canvas = tk.Canvas(tree_frame)
    scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=canvas.yview)
    frame = ttk.Frame(canvas)
    canvas.create_window((0, 0), window=frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    def on_frame_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))
    frame.bind("<Configure>", on_frame_configure)
    transposed_data = list(zip(*data))
    
    for i, column in enumerate(columns):
        bg_color = "#c7e3fc" if i % 2 == 0 else "#ffffff"
        field_label = tk.Label(frame, text=column, font=("Arial", 10, "bold"), bg=bg_color, anchor="w")
        field_label.grid(row=i, column=0, sticky="w", padx=5, pady=2)
        
        if i < len(transposed_data):
            for j, value in enumerate(transposed_data[i]):
                value_entry = tk.Entry(frame, font=("Arial", 10), bg=bg_color, readonlybackground=bg_color, bd=0, relief="flat")
                value_entry.insert(0, value)
                value_entry.config(state="readonly")
                value_entry.grid(row=i, column=j+1, sticky="ew", padx=5, pady=2)
                frame.columnconfigure(j+1, weight=1)
                value_entry.update_idletasks()
                text_width = value_entry.winfo_reqwidth()
                value_entry.config(width=text_width)

def display_admin_table():
    """Display 'AdminSystemInfo' table data with columns as rows, alternating row colors, and adjustable second column width."""
    columns, data = fetch_admin_data()
    for widget in admin_data_frame.winfo_children():
        widget.destroy()
    canvas = tk.Canvas(admin_data_frame)
    scrollbar = ttk.Scrollbar(admin_data_frame, orient="vertical", command=canvas.yview)
    frame = ttk.Frame(canvas)
    canvas.create_window((0, 0), window=frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    def on_frame_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))
    frame.bind("<Configure>", on_frame_configure)
    transposed_data = list(zip(*data))
    
    for i, column in enumerate(columns):
        bg_color = "#c7e3fc" if i % 2 == 0 else "#ffffff"
        field_label = tk.Label(frame, text=column, font=("Arial", 10, "bold"), bg=bg_color, anchor="w")
        field_label.grid(row=i, column=0, sticky="w", padx=5, pady=2)
        
        if i < len(transposed_data):
            for j, value in enumerate(transposed_data[i]):
                if column == "Drivers":
                    value = value.replace(") Name:", ")\nName:")
                value_entry = tk.Text(frame, font=("Arial", 10), bg=bg_color, bd=0, relief="flat", wrap="word", height=5)
                value_entry.insert("1.0", value)
                value_entry.config(state="disabled")
                value_entry.grid(row=i, column=j+1, sticky="ew", padx=5, pady=2)
                frame.columnconfigure(j+1, weight=1)
                value_entry.update_idletasks()
                text_width = value_entry.winfo_reqwidth()
                value_entry.config(width=text_width)

def update_database():
    """Update the database path when a new database is selected."""
    global db_path
    selected_db = db_var.get()
    if not selected_db.endswith(".db"):
        selected_db += ".db"
    if selected_db:
        db_path = os.path.join(db_directory, selected_db)
        if check_SystemInfo_exists():
            display_table()
        if check_admin_table_exists():
            display_admin_table()

def refresh_databases():
    """Run PowerShell script, refresh the dropdown list with available databases, and auto-select ComputerName.db if it exists."""
    sys.stdout.write("Running system audit...\n")
    try:
        subprocess.run(["powershell.exe", "-ExecutionPolicy", "Bypass", "-File", audit_directory], check=True, creationflags=subprocess.CREATE_NO_WINDOW)
    except subprocess.CalledProcessError as e:
        messagebox.showerror("PowerShell Error", f"An error occurred while running the PowerShell script: {e}")
        return
    available_dbs = fetch_databases()
    db_names = [os.path.splitext(db)[0] for db in available_dbs]
    db_dropdown["values"] = db_names
    computer_name = socket.gethostname()
    if computer_name in db_names:
        db_var.set(computer_name)
    elif db_names:
        db_var.set(db_names[0])
    sys.stdout.write("Audit Complete.\n")
    update_database()

def run_utility(path, admin):
    try:
        if path.lower().endswith('.vbs'):
            if admin == "1":
                ctypes.windll.shell32.ShellExecuteW(None, "runas", "cscript.exe", path, None, 1)
            else:
                subprocess.run(["cscript.exe", path], check=True)
        else:
            if admin == "1":
                ctypes.windll.shell32.ShellExecuteW(None, "runas", path, None, None, 1)
            else:
                subprocess.run([path], check=True)
    except Exception as e:
        messagebox.showerror("Execution Error", str(e))
        
def create_utility_buttons():
    school_drive = get_school_drive()
    for i in range(1, 25):
        section = f"Utilities{i}"
        name = config.get(section, "Name", fallback="")
        path = config.get(section, "Path", fallback="")
        admin = config.get(section, "Admin", fallback="0")
        if path.startswith("H:\\"):
            username = get_current_username()
            path = path.replace("H:\\", f"\\\\ad.ccrsb.ca\\it-home\\IT-SCHOOL-HOME\\{username}\\", 1)
        elif path.startswith("X:\\"):
            path = path.replace("X:\\", school_drive, 1)
        if name:
            button = tk.Button(utilities_frame, text=name, width=20, height=3, command=lambda p=path, a=admin: run_utility(p, a))
            button.grid(row=(i-1)//4, column=(i-1)%4, padx=10, pady=10, sticky="ew")

def export_data():
    """Run the Audit-Export.ps1 script with a file dialog box for the location of the xlsx file."""
    default_name = "SystemInfo.xlsx"
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=default_name, filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if file_path:
        subprocess.run(["powershell.exe", "-ExecutionPolicy", "Bypass", "-File", export_script_path, "-outputExcel", file_path], check=True, creationflags=subprocess.CREATE_NO_WINDOW, shell=False)
        if file_path:
            subprocess.run(["powershell.exe", "-ExecutionPolicy", "Bypass", "-File", export_script_path, "-outputExcel", file_path], check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            messagebox.showinfo("Export Complete", f"System information exported successfully to {file_path}")

def run_admin_audit():
    """Run the Admin-Audit.ps1 script as admin and update the Admin Info tab."""
    try:
        if not db_directory:
            messagebox.showerror("Error", "Database directory is not set.")
            return
        subprocess.run(["powershell.exe", "-ExecutionPolicy", "Bypass", "-File", admin_audit_script_path], check=True, creationflags=subprocess.CREATE_NO_WINDOW)
        display_admin_table()
    except Exception as e:
        messagebox.showerror("Execution Error", str(e))

def find_html_file():
    """Find the HTML file in the logs directory."""
    for file in os.listdir(logs_directory):
        if file.endswith(".html"):
            return os.path.join(logs_directory, file)
    return None

def open_html_file():
    """Open the HTML file in the default web browser."""
    html_file = find_html_file()
    if html_file:
        webbrowser.open(html_file)
    else:
        messagebox.showerror("File Not Found", "No HTML file found in the logs directory.")


# GUI Setup
root = tk.Tk()
root.title("CCRCE Login App")
root.geometry("800x500")
root.configure(bg="#5da9f0")
notebook = ttk.Notebook(root)
notebook.pack(expand=True, fill="both")
system_info_frame = ttk.Frame(notebook)
utilities_frame = ttk.Frame(notebook)
admin_frame = ttk.Frame(notebook)
notebook.add(system_info_frame, text="System Info")
notebook.add(utilities_frame, text="Utilities")
notebook.add(admin_frame, text="Admin Info")
control_frame = ttk.Frame(system_info_frame)
control_frame.pack(fill="x", padx=10, pady=5)
db_var = tk.StringVar()
db_dropdown = ttk.Combobox(control_frame, textvariable=db_var)
db_dropdown.pack(side="left", padx=5)
db_dropdown["values"] = fetch_databases()
db_dropdown.bind("<<ComboboxSelected>>", lambda e: update_database())
tk.Button(control_frame, text="Refresh", command=refresh_databases).pack(side="left", padx=5)
tk.Button(control_frame, text="Export", command=export_data).pack(side="left", padx=5)

# Admin Info tab layout
admin_button_frame = ttk.Frame(admin_frame)
admin_button_frame.pack(side="top", fill="x", padx=10, pady=5)
tk.Button(admin_button_frame, text="Update", command=run_admin_audit).pack(side="left", padx=5, pady=5)
tk.Button(admin_button_frame, text="Open Driver Report", command=open_html_file).pack(side="left", padx=5, pady=5)
admin_data_frame = ttk.Frame(admin_frame)
admin_data_frame.pack(expand=True, fill="both", padx=10, pady=5)

tree_frame = ttk.Frame(system_info_frame)
tree_frame.pack(expand=True, fill="both", padx=10, pady=5)

create_utility_buttons()
refresh_databases()
sys.stdout.write("Loading Complete.\n")
ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
root.mainloop()