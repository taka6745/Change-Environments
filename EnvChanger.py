import tkinter as tk
from tkinter import ttk, messagebox
import yaml
import os
import subprocess
import time

# Path to customers folder
CUSTOMERS_DIR = "customers"


def load_yaml(file_path):
    """Load YAML file."""
    with open(file_path, "r", encoding="utf-8") as file:
        return yaml.safe_load(file)


def get_prerequisite_versions():
    """Get installed prerequisite versions from both 32-bit and 64-bit registries."""
    try:
        ps_commands = [
            "Get-ItemProperty HKLM:\\Software\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\* "
            "| Where-Object { $_.DisplayName -like '*InformCAD Prerequisites*' } "
            "| Select-Object DisplayName, DisplayVersion",
            "Get-ItemProperty HKLM:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\* "
            "| Where-Object { $_.DisplayName -like '*InformCAD Prerequisites*' } "
            "| Select-Object DisplayName, DisplayVersion"
        ]
        results = []
        for ps_command in ps_commands:
            result = subprocess.check_output(["powershell", "-Command", ps_command], text=True)
            lines = result.strip().splitlines()
            if len(lines) > 2:  # Ensure there is data beyond headers
                results.append("\n".join(lines[2:]))  # Skip headers (DisplayName and DisplayVersion)

        if results:
            installed_version = results[0].split(":")[-1].strip()  # Extract version
            return installed_version
        return "0.0.0.0"  # Default if not found
    except subprocess.CalledProcessError:
        return "0.0.0.0"  # Default if an error occurs


def parse_prerequisite_version(version_str):
    """Extract the major version number from the version string."""
    try:
        return int(version_str.split('.')[0])  # Extract the major version (e.g., 23 from 23.1.3.3)
    except ValueError:
        return None


def uninstall_prerequisites():
    """Uninstall current prerequisites if found."""
    try:
        ps_command = (
            "Get-ItemProperty HKLM:\\Software\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\* "
            "| Where-Object { $_.DisplayName -like '*InformCAD Prerequisites*' } "
            "| Select-Object UninstallString"
        )
        uninstall_path = subprocess.check_output(["powershell", "-Command", ps_command], text=True).strip()
        lines = uninstall_path.splitlines()
        if len(lines) > 1:  # Ensure data beyond headers
            uninstall_exe = lines[-1].strip()
            messagebox.showinfo("Uninstalling Prerequisites", "Uninstaller starting")
            subprocess.run(uninstall_exe, shell=True)
        else:
            messagebox.showwarning("Uninstaller Not Found", "No existing prerequisites found to uninstall.")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"Failed to start uninstaller: {e}")


def install_prerequisites(required_version):
    """Install the correct prerequisites based on the required version."""
    major_version = parse_prerequisite_version(required_version)
    if major_version is None:
        messagebox.showerror("Error", "Invalid version format in YAML.")
        return

    if major_version < 20:
        prereq_path = r"Y:\Software\Suppliers\TriTech\software\Command_5.8.x\5.8.22 InformCAD Prerequisite\InformCAD Prerequisites.msi"
    else:
        prereq_path = r"Y:\Software\Suppliers\TriTech\software\CAD Enterprise 22\22.3.1_Prereq\22.3.1_Prereq\InformCAD Prerequisites.msi"

    try:
        messagebox.showinfo("Installing Prerequisites", f"Installer Started: {prereq_path}")
        subprocess.run(f"msiexec /i \"{prereq_path}\"", shell=True)
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"Failed to start installer: {e}")


def ensure_correct_prerequisites(required_version):
    """Ensure the correct prerequisites are installed."""
    while True:
        installed_version = get_prerequisite_versions().replace("InformCAD Prerequisites ", "")
        if check_version_compatibility(required_version, installed_version):
            break
        else:
            messagebox.showinfo("Updating Prerequisites", "Incorrect prerequisites found. Updating...")
            uninstall_prerequisites()
            time.sleep(5)  # Wait before checking again
            install_prerequisites(required_version)
            time.sleep(10)  # Wait for installation to complete


def check_version_compatibility(required_version, installed_version):
    """Check if the installed prerequisite version is compatible with the required version."""
    required_major = parse_prerequisite_version(required_version)
    installed_major = parse_prerequisite_version(installed_version)

    if required_major is None or installed_major is None:
        return False  # Invalid version format

    if (installed_major >= 20 and required_major >= 20) or (installed_major < 20 and required_major < 20):
        return True  # Compatible
    return False


def on_change():
    """Handle change button click."""
    selected_customer = customer_var.get()
    selected_environment = environment_var.get()
    selected_type = type_var.get()

    if not (selected_customer and selected_environment and selected_type):
        messagebox.showerror("Error", "Please select all options!")
        return

    # Build the path for the script to run
    environment_folder = os.path.join(CUSTOMERS_DIR, selected_customer, selected_environment)
    customer_file = os.path.join(CUSTOMERS_DIR, selected_customer, f"{selected_customer}.yaml")
    data = load_yaml(customer_file)

    # Get the required version and script name
    try:
        required_version = next(
            item['Version']
            for item in data[selected_environment]
            if 'Version' in item
        )
        script_name = next(
            item[selected_type]
            for item in data[selected_environment]
            if selected_type in item
        )
    except StopIteration:
        messagebox.showerror("Error", f"Type '{selected_type}' or version not found!")
        return

    # Ensure correct prerequisites
    ensure_correct_prerequisites(required_version)

    # Construct full path for the script
    script_path = os.path.join(environment_folder, script_name)

    if os.path.exists(script_path):
        try:
            # Execute the script using PowerShell
            subprocess.run(["powershell", "-File", script_path], check=True)
            messagebox.showinfo("Success", f"Executed: {script_path}")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Error", f"Error running script: {e}")
    else:
        messagebox.showerror("Error", f"Script not found: {script_path}")


def populate_environment_dropdown(event):
    """Update environments when customer is selected."""
    selected_customer = customer_var.get()
    if not selected_customer:
        return

    customer_file = os.path.join(CUSTOMERS_DIR, selected_customer, f"{selected_customer}.yaml")
    data = load_yaml(customer_file)
    environments = list(data.keys())
    environment_dropdown["values"] = environments
    environment_var.set("")  # Clear previous selection


def populate_type_dropdown(event):
    """Update types when environment is selected."""
    selected_customer = customer_var.get()
    selected_environment = environment_var.get()

    if not (selected_customer and selected_environment):
        return

    customer_file = os.path.join(CUSTOMERS_DIR, selected_customer, f"{selected_customer}.yaml")
    data = load_yaml(customer_file)

    # Extract only the keys (e.g., AppService, Workstation), exclude 'Version'
    types = []
    for item in data[selected_environment]:
        for key in item.keys():
            if key != "Version":
                types.append(key)
    type_dropdown["values"] = types
    type_var.set("")  # Clear previous selection


# Main application
if __name__ == "__main__":
    # Fetch installed prerequisites
    prerequisites = get_prerequisite_versions()

    # Tkinter GUI
    root = tk.Tk()
    root.title("Customer Environment Selector")

    # Display prerequisites at the top
    prerequisites_label = tk.Label(
        root,
        text=f"Installed Prerequisites:\n{prerequisites}",
        justify="left",
        anchor="w",
        padx=10,
        pady=10,
    )
    prerequisites_label.grid(row=0, column=0, columnspan=2, sticky="w")

    # Dropdowns and buttons
    tk.Label(root, text="Customer:").grid(row=1, column=0, padx=10, pady=10)
    customer_var = tk.StringVar()
    customer_dropdown = ttk.Combobox(root, textvariable=customer_var, state="readonly")
    customer_dropdown.grid(row=1, column=1, padx=10, pady=10)

    tk.Label(root, text="Environment:").grid(row=2, column=0, padx=10, pady=10)
    environment_var = tk.StringVar()
    environment_dropdown = ttk.Combobox(root, textvariable=environment_var, state="readonly")
    environment_dropdown.grid(row=2, column=1, padx=10, pady=10)

    tk.Label(root, text="Type:").grid(row=3, column=0, padx=10, pady=10)
    type_var = tk.StringVar()
    type_dropdown = ttk.Combobox(root, textvariable=type_var, state="readonly")
    type_dropdown.grid(row=3, column=1, padx=10, pady=10)

    change_button = ttk.Button(root, text="Change", command=on_change)
    change_button.grid(row=4, column=0, columnspan=2, pady=20)

    # Populate customer dropdown
    customers_data = load_yaml(os.path.join(CUSTOMERS_DIR, "customers.yaml"))
    customer_dropdown["values"] = customers_data["customers"]

    # Event bindings
    customer_dropdown.bind("<<ComboboxSelected>>", populate_environment_dropdown)
    environment_dropdown.bind("<<ComboboxSelected>>", populate_type_dropdown)

    root.mainloop()
