import tkinter as tk
from tkinter import ttk, messagebox
import yaml
import os
import subprocess

##############################
#  STEALTH WINDOW FUNCTIONS  #
##############################
def on_enter(event):
    """Mouse enters the main window area."""
    # If user is not in the middle of using a dropdown, show fully.
    if not is_in_combobox_focus():
        root.attributes("-alpha", 1.0)

def on_leave(event):
    """Mouse leaves the main window area."""
    # If the focus is still on one of our comboboxes or its popup, do not hide yet.
    if not is_in_combobox_focus():
        root.attributes("-alpha", 0.01)

def is_in_combobox_focus():
    """
    Return True if focus is on any combobox or its popup menu.
    This avoids KeyError on 'popdown' and keeps the window visible while combos are open.
    """
    try:
        focus_widget = root.focus_get()
        if not focus_widget:
            return False
        
        # Convert the widget to a string name
        widget_name = str(focus_widget)
        # If the combobox has spawned a popup widget named something like '.!combobox.popdown.focustrap',
        # we'll consider that as "in combobox focus" so we remain visible.
        if "popdown" in widget_name.lower():
            return True

        # Attempt to map the widget name back to an actual widget object
        w = root.nametowidget(widget_name)
        return (w in (customer_dropdown, environment_dropdown, type_dropdown))
    except (KeyError, tk.TclError):
        # If we cannot find the widget, assume it's a combobox popup or something related
        return True

def on_combobox_click(event):
    """When user clicks a combobox, ensure window is fully visible."""
    root.attributes("-alpha", 1.0)

def on_combobox_unfocus(event):
    """When combobox loses focus, possibly hide if mouse is outside window and not in another combobox."""
    if not is_in_combobox_focus():
        # Simulate a <Leave> event to hide if we're truly outside
        _event = tk.Event()
        on_leave(_event)

##############################
#     WINDOW DRAGGING        #
##############################
def start_move(event):
    """Remember the mouse offset within the title bar for window dragging."""
    root._drag_x = event.x
    root._drag_y = event.y

def on_drag(event):
    """Drag (move) the entire window based on mouse movement in the title bar."""
    x = root.winfo_x() + (event.x - root._drag_x)
    y = root.winfo_y() + (event.y - root._drag_y)
    root.geometry(f"+{x}+{y}")

def close_app():
    """Close the application."""
    root.destroy()

##############################
#        POSITIONING         #
##############################
def move_to_top_left(root, window_width=300, window_height=180, x_offset=25, y_offset=25):
    """
    Position window near the top-left corner of the screen
    with a small offset so it’s not flush against the edges.
    """
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    x = 0 + x_offset
    y = 0 + y_offset
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")

##############################
#       YAML + LOGIC         #
##############################
def load_config():
    """Load configuration file."""
    with open("Config.yaml", "r", encoding="utf-8") as file:
        return yaml.safe_load(file)

def get_customer_environments(config, customer_name):
    """Get environments for a specific customer."""
    for customer in config['customers']:
        if customer['name'] == customer_name:
            return customer['environments']
    return []

def get_environment_version(config, customer_name, environment_name):
    """Get version for a specific environment."""
    environments = get_customer_environments(config, customer_name)
    for env in environments:
        if env['name'] == environment_name:
            return env['version']
    return None

def get_environment_attributes(config, customer_name, environment_name):
    """Get attributes for a specific environment."""
    environments = get_customer_environments(config, customer_name)
    for env in environments:
        if env['name'] == environment_name:
            return env['attributes']
    return {}

def on_change():
    """Handle change button click."""
    selected_customer = customer_var.get()
    selected_environment = environment_var.get()
    selected_type = type_var.get()

    if not (selected_customer and selected_environment and selected_type):
        messagebox.showerror("Error", "Please select all options!")
        return

    # Get version and attributes
    version = get_environment_version(config, selected_customer, selected_environment)
    attributes = get_environment_attributes(config, selected_customer, selected_environment)

    if not version:
        messagebox.showerror("Error", "Version not found!")
        return

    major_version = version.split('.')[0]
    script_name = f"{major_version}_{selected_type}.ps1"

    if os.path.exists(script_name):
        try:
            # Execute the script using PowerShell with parameters
            if selected_type == "Workstation":
                if major_version == "23":
                    if attributes.get('install_path', '') == "":
                        appservice = attributes.get('appservice', '')
                        ps_command = [
                        "powershell",
                        "-File", script_name,
                        "-CustomerName", selected_customer,
                        "-Environment", selected_environment,
                        "-AppService", appservice,
                    ]
                    else:
                        appservice = attributes.get('appservice', '')
                        install_path = attributes.get('install_path', '')

                        ps_command = [
                        "powershell",
                        "-File", script_name,
                        "-CustomerName", selected_customer,
                        "-Environment", selected_environment,
                        "-AppService", appservice,
                        "-InstallPath", install_path
                        ]
                elif major_version == "5":
                    qdrive = attributes.get('qdrive', '')

                    ps_command = [
                    "powershell",
                    "-File", script_name,
                    "-CustomerName", selected_customer,
                    "-Environment", selected_environment,
                    "-QDrive", qdrive
                    ]
            elif selected_type == "Interface":
                qdrive = attributes.get('qdrive', '')
                ps_command = [
                    "powershell",
                    "-File", script_name,
                    "-CustomerName", selected_customer,
                    "-Environment", selected_environment,
                    "-QDrive", qdrive
                ]
            elif selected_type == "Service":
                qdrive = attributes.get('qdrive', '')
                ps_command = [
                    "powershell",
                    "-File", script_name,
                    "-CustomerName", selected_customer,
                    "-Environment", selected_environment,
                    "-QDrive", qdrive
                ]
            subprocess.run(ps_command, check=True)
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Error", f"Error running script: {e}")
    else:
        messagebox.showerror("Error", f"Script not found: {script_name}")

def populate_environment_dropdown(event):
    """Update environments when customer is selected."""
    selected_customer = customer_var.get()
    if not selected_customer:
        return

    environments = [env['name'] for env in get_customer_environments(config, selected_customer)]
    environment_dropdown["values"] = environments
    environment_var.set("")

#################################
#           MAIN APP            #
#################################
if __name__ == "__main__":
    config = load_config()

    root = tk.Tk()
    root.title("Customer Environment Selector")

    # Overrideredirect removes the standard title bar
    root.overrideredirect(True)
    # Always on top
    root.wm_attributes("-topmost", True)
    # Start mostly invisible
    root.attributes("-alpha", 0.01)

    # Position in top-left corner, offset so user sees a small border
    move_to_top_left(root, window_width=280, window_height=200, x_offset=25, y_offset=25)

    # Bind Enter/Leave events to show/hide the entire window
    root.bind("<Enter>", on_enter)
    root.bind("<Leave>", on_leave)

    #
    # CUSTOM TITLE BAR
    #
    title_bar = tk.Frame(root, bg="gray", height=30)
    title_bar.pack(fill="x", side="top")

    # Close button
    close_btn = tk.Button(title_bar, text=" X ", command=close_app, bg="gray", fg="white", bd=0)
    close_btn.pack(side="right")

    # A little label to let user click+drag
    # (We could make the entire title_bar draggable, so we bind to it.)
    title_label = tk.Label(title_bar, text="  Customer Environment Selector  ",
                           bg="gray", fg="white")
    title_label.pack(side="left", padx=(5,0))

    # Bind the dragging events to the entire title_bar
    title_bar.bind("<Button-1>", start_move)
    title_bar.bind("<B1-Motion>", on_drag)

    #
    # MAIN FRAME
    #
    main_frame = tk.Frame(root, bg="lightgray")
    main_frame.pack(fill="both", expand=True)

    # UI Layout
    tk.Label(main_frame, text="Customer:", bg="lightgray").grid(row=0, column=0, padx=10, pady=10, sticky="e")
    customer_var = tk.StringVar()
    customer_dropdown = ttk.Combobox(main_frame, textvariable=customer_var, state="readonly")
    customer_dropdown["values"] = [customer['name'] for customer in config['customers']]
    customer_dropdown.grid(row=0, column=1, padx=10, pady=10)

    tk.Label(main_frame, text="Environment:", bg="lightgray").grid(row=1, column=0, padx=10, pady=10, sticky="e")
    environment_var = tk.StringVar()
    environment_dropdown = ttk.Combobox(main_frame, textvariable=environment_var, state="readonly")
    environment_dropdown.grid(row=1, column=1, padx=10, pady=10)

    tk.Label(main_frame, text="Type:", bg="lightgray").grid(row=2, column=0, padx=10, pady=10, sticky="e")
    type_var = tk.StringVar()
    type_dropdown = ttk.Combobox(main_frame, textvariable=type_var, state="readonly")
    type_dropdown["values"] = ["Workstation", "Interface", "Service"]
    type_dropdown.grid(row=2, column=1, padx=10, pady=10)

    change_button = ttk.Button(main_frame, text="Change", command=on_change)
    change_button.grid(row=3, column=0, columnspan=2, pady=20)

    # When a customer is selected, update environment dropdown
    customer_dropdown.bind("<<ComboboxSelected>>", populate_environment_dropdown)

    # Ensure the window remains visible if user opens a dropdown
    # (bind combobox focus events)
    for cb in (customer_dropdown, environment_dropdown, type_dropdown):
        cb.bind("<Button-1>", on_combobox_click)   # user clicks to open
        cb.bind("<FocusOut>", on_combobox_unfocus) # user stops focusing combobox

    root.mainloop()
