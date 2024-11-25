import tkinter as tk
from tkinter import ttk, messagebox
import yaml
import os
import subprocess


# Path to customers folder
CUSTOMERS_DIR = "customers"


def load_yaml(file_path):
    """Load YAML file."""
    with open(file_path, "r", encoding="utf-8") as file:
        return yaml.safe_load(file)


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

    # Get the script name
    try:
        script_name = next(
            item[selected_type]
            for item in data[selected_environment]
            if selected_type in item
        )
    except StopIteration:
        messagebox.showerror("Error", f"Type '{selected_type}' not found!")
        return

    # Construct full path
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

    # Extract only the keys (e.g., AppService, Workstation)
    types = []
    for item in data[selected_environment]:
        types.extend(item.keys())
    type_dropdown["values"] = types
    type_var.set("")  # Clear previous selection


# Main application
if __name__ == "__main__":
    # Tkinter GUI
    root = tk.Tk()
    root.title("Customer Environment Selector")

    tk.Label(root, text="Customer:").grid(row=0, column=0, padx=10, pady=10)
    customer_var = tk.StringVar()
    customer_dropdown = ttk.Combobox(root, textvariable=customer_var, state="readonly")
    customer_dropdown.grid(row=0, column=1, padx=10, pady=10)

    tk.Label(root, text="Environment:").grid(row=1, column=0, padx=10, pady=10)
    environment_var = tk.StringVar()
    environment_dropdown = ttk.Combobox(root, textvariable=environment_var, state="readonly")
    environment_dropdown.grid(row=1, column=1, padx=10, pady=10)

    tk.Label(root, text="Type:").grid(row=2, column=0, padx=10, pady=10)
    type_var = tk.StringVar()
    type_dropdown = ttk.Combobox(root, textvariable=type_var, state="readonly")
    type_dropdown.grid(row=2, column=1, padx=10, pady=10)

    change_button = ttk.Button(root, text="Change", command=on_change)
    change_button.grid(row=3, column=0, columnspan=2, pady=20)

    # Populate customer dropdown
    customers_data = load_yaml(os.path.join(CUSTOMERS_DIR, "customers.yaml"))
    customer_dropdown["values"] = customers_data["customers"]

    # Event bindings
    customer_dropdown.bind("<<ComboboxSelected>>", populate_environment_dropdown)
    environment_dropdown.bind("<<ComboboxSelected>>", populate_type_dropdown)

    root.mainloop()
