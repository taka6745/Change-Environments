import tkinter as tk
from tkinter import ttk, messagebox
import yaml
import os
import argparse


# Path to customers folder
CUSTOMERS_DIR = "customers"


def load_yaml(file_path):
    """Load YAML file."""
    with open(file_path, "r") as file:
        return yaml.safe_load(file)


def on_change():
    """Handle change button click."""
    selected_customer = customer_var.get()
    selected_environment = environment_var.get()
    selected_type = type_var.get()

    if not (selected_customer and selected_environment and selected_type):
        messagebox.showerror("Error", "Please select all options!")
        return

    customer_file = os.path.join(CUSTOMERS_DIR, selected_customer, f"{selected_customer}.yaml")
    data = load_yaml(customer_file)

    try:
        script_or_url = data[selected_environment][0][selected_type]
        messagebox.showinfo("Action", f"Execute: {script_or_url}")
    except KeyError:
        messagebox.showerror("Error", "Invalid selection or missing script/type!")


def populate_environment_dropdown(event):
    """Update environments when customer is selected."""
    selected_customer = customer_var.get()
    if not selected_customer:
        return

    customer_file = os.path.join(CUSTOMERS_DIR, selected_customer, f"{selected_customer}.yaml")
    data = load_yaml(customer_file)
    environments = data.keys()
    environment_dropdown["values"] = list(environments)


def populate_type_dropdown(event):
    """Update types when environment is selected."""
    selected_customer = customer_var.get()
    selected_environment = environment_var.get()
    if not (selected_customer and selected_environment):
        return

    customer_file = os.path.join(CUSTOMERS_DIR, selected_customer, f"{selected_customer}.yaml")
    data = load_yaml(customer_file)
    types = data[selected_environment][0].keys()
    type_dropdown["values"] = list(types)


# Main application
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Launch GUI or run CLI.")
    parser.add_argument("--cli", action="store_true", help="Run in CLI mode.")
    args = parser.parse_args()

    if args.cli:
        print("CLI mode is not implemented yet.")
        exit(0)

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
