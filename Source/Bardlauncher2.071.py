import os
import json
import shutil
import time
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk, Menu, simpledialog
from PIL import Image, ImageTk
from win32com.client import Dispatch
from idlelib.tooltip import Hovertip
from tkhtmlview import HTMLLabel
from ttkthemes import ThemedTk
import markdown
import psutil  # for checking if a process is running
from datetime import datetime

CONFIG_FILE = "bard_launcher_config.json"

class ScrollableFrame(ttk.Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        canvas = tk.Canvas(self)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

class BardLauncherGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Bard Launcher 2.07.2")
        self.root.geometry("800x600")

        # Track if the Start All or Start Selected buttons have been pressed
        self.start_all_pressed = False
        self.start_selected_pressed = False

        # Set the initial theme
        self.style = ttk.Style(self.root)
        self.style.theme_use('clam')  # Default light theme

        # Set the icon
        self.default_icon_path = os.path.join(os.getcwd(), "icon.png")  # Use the icon file created
        if os.path.exists(self.default_icon_path):
            self.icon = tk.PhotoImage(file=self.default_icon_path)
            self.root.iconphoto(True, self.icon)
        
        # Create a Notebook for tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.grid(row=0, column=0, columnspan=4, padx=5, pady=5, sticky="nsew")

        # Create frames for tabs
        self.main_frame = ttk.Frame(self.notebook)
        self.settings_frame = ttk.Frame(self.notebook)
        self.readme_frame = ttk.Frame(self.notebook)
        
        self.notebook.add(self.main_frame, text="Main")
        self.notebook.add(self.settings_frame, text="Settings")
        self.notebook.add(self.readme_frame, text="Readme")

        # Configure grid layout to make the frames expand
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        self.notebook.grid_rowconfigure(0, weight=1)
        self.notebook.grid_columnconfigure(0, weight=1)

        # Load saved config if it exists
        self.config_data = self.load_config()

        # Config Directory
        self.config_dir_label = ttk.Label(self.settings_frame, text="Config Directory")
        self.config_dir_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.config_dir_entry = ttk.Entry(self.settings_frame, width=50)
        self.config_dir_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.config_dir_button = ttk.Button(self.settings_frame, text="Browse", command=self.browse_config_dir)
        self.config_dir_button.grid(row=0, column=2, padx=5, pady=5)
        Hovertip(self.config_dir_button, 'Browse to select the configuration directory')

        # Shortcut Directory
        self.shortcut_dir_label = ttk.Label(self.settings_frame, text="Shortcut Directory")
        self.shortcut_dir_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.shortcut_dir_entry = ttk.Entry(self.settings_frame, width=50)
        self.shortcut_dir_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.shortcut_dir_button = ttk.Button(self.settings_frame, text="Browse", command=self.browse_shortcut_dir)
        self.shortcut_dir_button.grid(row=1, column=2, padx=5, pady=5)
        Hovertip(self.shortcut_dir_button, 'Browse to select the shortcut directory')

        # Delay Entry
        self.delay_label = ttk.Label(self.settings_frame, text="Seconds Delay")
        self.delay_label.grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.delay_entry = ttk.Entry(self.settings_frame, width=5)
        self.delay_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        self.delay_entry.insert(0, "10")  # Default value
        Hovertip(self.delay_entry, 'Set the delay in seconds between launching each shortcut')

        # Dark Mode Toggle
        self.dark_mode_var = tk.BooleanVar()
        self.dark_mode_checkbutton = ttk.Checkbutton(self.settings_frame, text="Dark Mode", variable=self.dark_mode_var, command=self.toggle_dark_mode)
        self.dark_mode_checkbutton.grid(row=4, column=1, padx=5, pady=5, sticky="w")
        Hovertip(self.dark_mode_checkbutton, 'Toggle dark mode')

        # Start All Button
        self.start_all_button = ttk.Button(self.main_frame, text="Start All", command=self.confirm_start_all_process)
        self.start_all_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        Hovertip(self.start_all_button, 'Start all shortcuts with the specified delay')

        # Start Selected Button
        self.start_selected_button = ttk.Button(self.main_frame, text="Start Selected", command=self.confirm_start_selected_process)
        self.start_selected_button.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        Hovertip(self.start_selected_button, 'Start only the selected shortcuts with the specified delay')

        # Additional Button
        self.additional_button = ttk.Button(self.main_frame, text="Move Default Config", command=self.move_default_config)
        self.additional_button.grid(row=0, column=2, padx=5, pady=5, sticky="ew")
        Hovertip(self.additional_button, 'Move the default configuration file to the selected config directory')

        # Toggle View Mode
        self.view_mode_var = tk.BooleanVar()
        self.view_mode_checkbutton = ttk.Checkbutton(self.main_frame, text="Grid View", variable=self.view_mode_var, command=self.populate_shortcuts)
        self.view_mode_checkbutton.grid(row=0, column=3, padx=5, pady=5, sticky="e")
        Hovertip(self.view_mode_checkbutton, 'Toggle between list and grid view')

        # Bard Buttons
        self.bard_buttons_frame = ttk.Frame(self.main_frame)
        self.bard_buttons_frame.grid(row=1, column=0, columnspan=4, padx=5, pady=5, sticky="nsew")

        # Progress Bar
        self.progress_bar = ttk.Progressbar(self.main_frame, orient="horizontal", length=400, mode="determinate")
        self.progress_bar.grid(row=2, column=0, columnspan=4, padx=5, pady=5, sticky="ew")

        # Status Display
        self.status_label = ttk.Label(self.main_frame, text="Status")
        self.status_label.grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.status_text = scrolledtext.ScrolledText(self.main_frame, height=10, width=80)
        self.status_text.grid(row=4, column=0, columnspan=4, padx=5, pady=5, sticky="nsew")

        # Clear Status Button
        self.clear_status_button = ttk.Button(self.main_frame, text="Clear Status", command=self.clear_status)
        self.clear_status_button.grid(row=5, column=0, padx=5, pady=5, sticky="ew")
        Hovertip(self.clear_status_button, 'Clear the status log')

        # Save Settings Button
        self.save_settings_button = ttk.Button(self.settings_frame, text="Save Settings", command=self.save_settings)
        self.save_settings_button.grid(row=3, column=0, padx=5, pady=5, sticky="ew")
        Hovertip(self.save_settings_button, 'Save the current settings')

        # Load Settings Button
        self.load_settings_button = ttk.Button(self.settings_frame, text="Load Settings", command=self.load_settings)
        self.load_settings_button.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
        Hovertip(self.load_settings_button, 'Load the previously saved settings')

        # Reset Configuration Button
        self.reset_config_button = ttk.Button(self.settings_frame, text="Reset Configuration", command=self.reset_configuration)
        self.reset_config_button.grid(row=3, column=2, padx=5, pady=5, sticky="ew")
        Hovertip(self.reset_config_button, 'Reset the configuration to default settings')

        # Readme Tab
        self.readme_text = HTMLLabel(self.readme_frame, html=self.load_readme())
        self.readme_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.bard_checkbuttons = {}

        # Experimental Section
        self.experimental_section = ttk.LabelFrame(self.settings_frame, text="Experimental")
        self.experimental_section.grid(row=5, column=0, columnspan=3, padx=5, pady=5, sticky="nsew")

        self.experimental_section.grid_rowconfigure(0, weight=1)
        self.experimental_section.grid_columnconfigure(0, weight=1)
        self.experimental_vars = {}

        separator = ttk.Separator(self.experimental_section, orient='horizontal')
        separator.grid(row=0, column=0, columnspan=3, padx=5, pady=10, sticky="ew")

        # Run LightAmp
        self.lightamp_check_var = tk.BooleanVar()
        self.lightamp_checkbutton = ttk.Checkbutton(self.experimental_section, text="Run LightAmp", variable=self.lightamp_check_var)
        self.lightamp_checkbutton.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        Hovertip(self.lightamp_checkbutton, 'Enable or disable running LightAmp before launching bards')

        self.lightamp_label = ttk.Label(self.experimental_section, text="LightAmp Location:")
        self.lightamp_label.grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.lightamp_entry = ttk.Entry(self.experimental_section, width=50)
        self.lightamp_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        self.lightamp_browse_button = ttk.Button(self.experimental_section, text="Browse", command=self.browse_lightamp)
        self.lightamp_browse_button.grid(row=2, column=2, padx=5, pady=5)
        Hovertip(self.lightamp_browse_button, 'Browse to select the LightAmp executable')

        # Separator
        separator = ttk.Separator(self.experimental_section, orient='horizontal')
        separator.grid(row=3, column=0, columnspan=3, padx=5, pady=10, sticky="ew")

        # Shortcut Creator Header
        self.shortcut_creator_label = ttk.Label(self.experimental_section, text="Shortcut Creator", font=('Helvetica', 12, 'bold'))
        self.shortcut_creator_label.grid(row=4, column=0, columnspan=3, padx=5, pady=5)

        self.json_label = ttk.Label(self.experimental_section, text="accountsList.json Path:")
        self.json_label.grid(row=5, column=0, padx=5, pady=5, sticky="e")
        self.json_entry = ttk.Entry(self.experimental_section, width=50)
        self.json_entry.grid(row=5, column=1, padx=5, pady=5, sticky="ew")
        self.json_browse_button = ttk.Button(self.experimental_section, text="Browse", command=self.browse_json)
        self.json_browse_button.grid(row=5, column=2, padx=5, pady=5)
        Hovertip(self.json_browse_button, 'Browse to select the accountsList.json file')

        self.shortcut_label = ttk.Label(self.experimental_section, text="Shortcut Directory:")
        self.shortcut_label.grid(row=6, column=0, padx=5, pady=5, sticky="e")
        self.shortcut_entry = ttk.Entry(self.experimental_section, width=50)
        self.shortcut_entry.grid(row=6, column=1, padx=5, pady=5, sticky="ew")
        self.shortcut_browse_button = ttk.Button(self.experimental_section, text="Browse", command=self.browse_shortcut)
        self.shortcut_browse_button.grid(row=6, column=2, padx=5, pady=5)
        Hovertip(self.shortcut_browse_button, 'Browse to select the shortcut directory')

        self.roaming_check_var = tk.BooleanVar()
        self.roaming_checkbutton = ttk.Checkbutton(self.experimental_section, text="Use Roaming Directory", variable=self.roaming_check_var, command=self.toggle_roaming_path)
        self.roaming_checkbutton.grid(row=7, column=0, columnspan=3, padx=5, pady=5, sticky="w")
        Hovertip(self.roaming_checkbutton, 'Enable or disable the use of a roaming directory')

        self.roaming_label = ttk.Label(self.experimental_section, text="Roaming Directory:")
        self.roaming_label.grid(row=8, column=0, padx=5, pady=5, sticky="e")
        self.roaming_entry = ttk.Entry(self.experimental_section, width=50, state="disabled")
        self.roaming_entry.grid(row=8, column=1, padx=5, pady=5, sticky="ew")
        self.roaming_browse_button = ttk.Button(self.experimental_section, text="Browse", command=self.browse_roaming, state="disabled")
        self.roaming_browse_button.grid(row=8, column=2, padx=5, pady=5)
        Hovertip(self.roaming_browse_button, 'Browse to select the roaming directory')

        self.create_button = ttk.Button(self.experimental_section, text="Create Shortcuts", command=self.create_shortcuts)
        self.create_button.grid(row=9, column=0, columnspan=3, padx=5, pady=10)
        Hovertip(self.create_button, 'Create shortcuts for the selected accounts')

        self.accounts_frame = ScrollableFrame(self.experimental_section)
        self.accounts_frame.grid(row=10, column=0, columnspan=3, padx=5, pady=5, sticky="nsew")
        self.accounts_vars = {}

        # Load previous paths and checkbox states
        if self.config_data:
            self.config_dir_entry.insert(0, self.config_data.get('config_dir', ''))
            self.shortcut_dir_entry.insert(0, self.config_data.get('shortcut_dir', ''))
            self.delay_entry.delete(0, tk.END)  # Clear the delay entry field before inserting new value
            self.delay_entry.insert(0, str(self.config_data.get('delay', 10)))
            self.dark_mode_var.set(self.config_data.get('dark_mode', False))
            self.toggle_dark_mode()  # Set initial theme based on saved config
            if 'bard_checkbuttons' in self.config_data:
                self.populate_shortcuts(self.config_data['bard_checkbuttons'])
            else:
                self.populate_shortcuts()
            self.lightamp_check_var.set(self.config_data.get('lightamp_check', False))
            self.lightamp_entry.insert(0, self.config_data.get('lightamp_location', ''))
        else:
            # Set a common default path for the config directory
            default_config_path = os.path.join(os.path.expanduser('~'), "Documents", "My Games", "FINAL FANTASY XIV - A Realm Reborn")
            if os.path.isdir(default_config_path):
                self.config_dir_entry.insert(0, default_config_path)

    def toggle_roaming_path(self):
        if self.roaming_check_var.get():
            self.roaming_entry.config(state="normal")
            self.roaming_browse_button.config(state="normal")
        else:
            self.roaming_entry.config(state="disabled")
            self.roaming_browse_button.config(state="disabled")

    def browse_json(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
        if file_path:
            self.json_entry.delete(0, tk.END)
            self.json_entry.insert(0, file_path)
            self.load_accounts(file_path)

    def browse_shortcut(self):
        directory = filedialog.askdirectory()
        if directory:
            self.shortcut_entry.delete(0, tk.END)
            self.shortcut_entry.insert(0, directory)

    def browse_roaming(self):
        directory = filedialog.askdirectory()
        if directory:
            self.roaming_entry.delete(0, tk.END)
            self.roaming_entry.insert(0, directory)

    def browse_lightamp(self):
        file_path = filedialog.askopenfilename(filetypes=[("Executable files", "*.exe"), ("All files", "*.*")])
        if file_path:
            self.lightamp_entry.delete(0, tk.END)
            self.lightamp_entry.insert(0, file_path)

    def load_accounts(self, file_path):
        with open(file_path, 'r') as file:
            accounts = json.load(file)

        for widget in self.accounts_frame.scrollable_frame.winfo_children():
            widget.destroy()

        self.accounts_vars = {}
        for account in accounts:
            user_name = account.get('UserName')
            var = tk.BooleanVar()
            chk = ttk.Checkbutton(self.accounts_frame.scrollable_frame, text=user_name, variable=var)
            chk.pack(anchor='w')
            self.accounts_vars[user_name] = var

    def create_shortcuts(self):
        json_file_path = self.json_entry.get()
        shortcut_directory = self.shortcut_entry.get()
        roaming_directory = self.roaming_entry.get() if self.roaming_check_var.get() else ""

        if not os.path.isfile(json_file_path):
            messagebox.showerror("Error", "Invalid accountsList.json file path")
            return

        if not os.path.exists(shortcut_directory):
            os.makedirs(shortcut_directory)

        with open(json_file_path, 'r') as file:
            accounts = json.load(file)

        xiv_launcher_path = os.path.expanduser(r"~\AppData\Local\XIVLauncher\XIVLauncher.exe")

        for account in accounts:
            user_name = account.get('UserName')
            if not self.accounts_vars[user_name].get():
                continue

            use_otp = account.get('UseOtp', False)
            use_steam = account.get('UseSteamServiceAccount', False)

            # Construct the arguments for the target
            args = f'--account={user_name}-{use_otp}-{use_steam}'
            if roaming_directory:
                args += f' --roamingPath={roaming_directory}'

            # Shortcut file name
            shortcut_name = f"{user_name}.lnk"
            shortcut_path = os.path.join(shortcut_directory, shortcut_name)

            # Create the shortcut
            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.TargetPath = xiv_launcher_path
            shortcut.Arguments = args
            shortcut.WorkingDirectory = os.path.dirname(xiv_launcher_path)
            shortcut.IconLocation = xiv_launcher_path
            shortcut.save()

            self.status_text.insert(tk.END, f"Created shortcut for {user_name} at {shortcut_path}\n")

    def load_readme(self):
        try:
            with open("Readme.txt", "r") as file:
                content = file.read()
                return markdown.markdown(content)
        except FileNotFoundError:
            return "<h2>Readme.txt file not found in the current directory.</h2>"

    def browse_config_dir(self):
        directory = filedialog.askdirectory()
        self.config_dir_entry.delete(0, tk.END)
        self.config_dir_entry.insert(0, directory)

    def browse_shortcut_dir(self):
        directory = filedialog.askdirectory()
        self.shortcut_dir_entry.delete(0, tk.END)
        self.shortcut_dir_entry.insert(0, directory)
        self.populate_shortcuts()

    def populate_shortcuts(self, checkbutton_states=None):
        shortcut_dir = self.shortcut_dir_entry.get()
        if not shortcut_dir:
            return

        for widget in self.bard_buttons_frame.winfo_children():
            widget.destroy()

        self.bard_checkbuttons = {}
        if self.view_mode_var.get():
            # Grid view
            row = 0
            col = 0
            max_cols = 4
            for file in os.listdir(shortcut_dir):
                bard_name = file.split(".")[0]
                var = tk.BooleanVar()
                if checkbutton_states and bard_name in checkbutton_states:
                    var.set(checkbutton_states[bard_name])
                self.bard_checkbuttons[bard_name] = var

                # Load the icon for the shortcut
                shortcut_path = os.path.join(shortcut_dir, file)
                icon_path = self.get_icon_path(shortcut_path)
                if not icon_path or not os.path.exists(icon_path):
                    icon_path = self.default_icon_path  # Use default icon if specific icon is missing
                icon_image = Image.open(icon_path)
                icon_image = icon_image.resize((64, 64), Image.LANCZOS)
                icon_photo = ImageTk.PhotoImage(icon_image)

                bard_button = ttk.Checkbutton(self.bard_buttons_frame, image=icon_photo, text=bard_name, variable=var, compound='top')
                bard_button.image = icon_photo  # Keep a reference to avoid garbage collection
                bard_button.grid(row=row, column=col, padx=10, pady=10)
                Hovertip(bard_button, bard_name)

                # Bind right-click to show context menu
                bard_button.bind("<Button-3>", lambda event, name=bard_name: self.show_context_menu(event, name))

                col += 1
                if col >= max_cols:
                    col = 0
                    row += 1
        else:
            # List view
            row = 0
            for file in os.listdir(shortcut_dir):
                bard_name = file.split(".")[0]
                var = tk.BooleanVar()
                if checkbutton_states and bard_name in checkbutton_states:
                    var.set(checkbutton_states[bard_name])
                checkbutton = ttk.Checkbutton(self.bard_buttons_frame, text=bard_name, variable=var)
                checkbutton.grid(row=row, column=0, padx=5, pady=5, sticky="w")
                self.bard_checkbuttons[bard_name] = var

                bard_button = ttk.Button(self.bard_buttons_frame, text=f"Launch {bard_name}", command=lambda name=bard_name: self.launch_bard(self.shortcut_dir_entry.get(), name))
                bard_button.grid(row=row, column=1, padx=5, pady=5)
                Hovertip(bard_button, f'Launch the shortcut for {bard_name}')
                copy_button = ttk.Button(self.bard_buttons_frame, text=f"Copy Config for {bard_name}", command=lambda name=bard_name: self.copy_config(name))
                copy_button.grid(row=row, column=2, padx=5, pady=5)
                Hovertip(copy_button, f'Copy the config for {bard_name}')
                row += 1

    def get_icon_path(self, shortcut_path):
        try:
            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(shortcut_path)
            icon_path = shortcut.IconLocation.split(',')[0]
            return icon_path
        except Exception as e:
            self.status_text.insert(tk.END, f"Error retrieving icon: {e}\n")
            return ''

    def toggle_dark_mode(self):
        if self.dark_mode_var.get():
            self.style.theme_use('black')  # Set dark theme
        else:
            self.style.theme_use('clam')  # Set light theme

    def confirm_start_all_process(self):
        if self.start_all_pressed:
            if not messagebox.askokcancel("Confirm", "The 'Start All' button has already been pressed. Do you want to run it again?"):
                return
        self.start_all_pressed = True
        self.start_all_process()

    def confirm_start_selected_process(self):
        if self.start_selected_pressed:
            if not messagebox.askokcancel("Confirm", "The 'Start Selected' button has already been pressed. Do you want to run it again?"):
                return
        self.start_selected_pressed = True
        self.start_selected_process()

    def start_all_process(self):
        self.start_process(selected_only=False)
        self.create_dynamic_buttons()

    def start_selected_process(self):
        self.start_process(selected_only=True)

    def create_dynamic_buttons(self):
        self.populate_shortcuts()

    def start_process(self, selected_only):
        config_dir = self.config_dir_entry.get()
        shortcut_dir = self.shortcut_dir_entry.get()
        try:
            delay = max(10, int(self.delay_entry.get()))
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid number for the delay.")
            return

        if not config_dir or not shortcut_dir:
            messagebox.showerror("Error", "Please select both directories.")
            return

        # Save paths and checkbox states to config file
        self.save_config(config_dir, shortcut_dir, delay, self.dark_mode_var.get(), {k: v.get() for k, v in self.bard_checkbuttons.items()}, self.lightamp_check_var.get(), self.lightamp_entry.get())

        config_file = os.path.join(os.path.expanduser('~'), "Documents", "My Games", "FINAL FANTASY XIV - A Realm Reborn", "FFXIV.cfg")
        self.status_text.insert(tk.END, "Starting process...\n")

        selected_bards = [bard_name for bard_name, var in self.bard_checkbuttons.items() if not selected_only or var.get()]
        self.progress_bar["maximum"] = len(selected_bards)

        if self.lightamp_check_var.get():
            self.start_lightamp()

        for i, bard_name in enumerate(selected_bards, start=1):
            self.status_text.insert(tk.END, f"Working on Bard {bard_name}.\n")

            config_file_path = os.path.join(config_dir, f"{bard_name}.cfg")
            if os.path.isfile(config_file_path):
                self.status_text.insert(tk.END, f"Found a config file for {bard_name}. Copying to FFXIV config {config_file}.\n")
                shutil.copy2(config_file_path, config_file)
                self.status_text.insert(tk.END, "Done.\n")
            else:
                self.status_text.insert(tk.END, f"Did not find a config file at {config_file_path} for {bard_name}\n")

            self.launch_bard(shortcut_dir, bard_name)
            self.progress_bar["value"] = i
            self.root.update_idletasks()

        self.status_text.insert(tk.END, "All Done!\n")

    def start_lightamp(self):
        # Check if LightAmp is already running
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] == 'LightAmp.exe':
                self.status_text.insert(tk.END, "LightAmp is already running.\n")
                return

        # Start LightAmp
        lightamp_path = self.lightamp_entry.get()
        if not os.path.isfile(lightamp_path):
            messagebox.showerror("Error", "Invalid LightAmp.exe file path")
            return

        try:
            os.startfile(lightamp_path)
            self.status_text.insert(tk.END, "Started LightAmp.\n")
        except Exception as e:
            self.status_text.insert(tk.END, f"Failed to start LightAmp. Error: {e}\n")

    def launch_bard(self, shortcut_dir, bard_name):
        self.status_text.insert(tk.END, f"Starting FFXIV for {bard_name}...\n")
        shortcut_path = os.path.join(shortcut_dir, f"{bard_name}.lnk")
        
        if not self.is_valid_xivlauncher_shortcut(shortcut_path):
            self.status_text.insert(tk.END, f"Invalid shortcut for {bard_name}. Skipping.\n")
            return
        
        try:
            os.startfile(shortcut_path)
        except Exception as e:
            self.status_text.insert(tk.END, f"Failed to launch the shortcut for {bard_name}. Error: {e}\n")
        else:
            self.status_text.insert(tk.END, f"Successfully launched FFXIV for {bard_name}.\n")
        self.root.update()
        time.sleep(max(10, int(self.delay_entry.get())))

    def is_valid_xivlauncher_shortcut(self, shortcut_path):
        try:
            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(shortcut_path)
            target_path = shortcut.TargetPath
            # Replace with the actual path to XIVLauncher.exe on your system
            expected_path = os.path.expanduser(r"~\AppData\Local\XIVLauncher\XIVLauncher.exe")
            return os.path.samefile(target_path, expected_path)
        except Exception as e:
            self.status_text.insert(tk.END, f"Error verifying shortcut: {e}\n")
            return False

    def copy_config(self, bard_name):
        config_dir = self.config_dir_entry.get()
        if not config_dir:
            messagebox.showerror("Error", "Please select the config directory.")
            return
        
        config_file = os.path.join(os.path.expanduser('~'), "Documents", "My Games", "FINAL FANTASY XIV - A Realm Reborn", "FFXIV.cfg")
        new_config_file_path = os.path.join(config_dir, f"{bard_name}.cfg")
        
        if messagebox.askokcancel("Confirm Copy", f"Are you sure you want to copy the config for {bard_name}?"):
            backup_dir = os.path.join(config_dir, "backup")
            os.makedirs(backup_dir, exist_ok=True)
            if os.path.isfile(new_config_file_path):
                backup_file_path = os.path.join(backup_dir, f"{bard_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.cfg")
                shutil.move(new_config_file_path, backup_file_path)
                self.status_text.insert(tk.END, f"Moved existing config to backup: {backup_file_path}\n")
            
            self.status_text.insert(tk.END, f"Copying config file to {new_config_file_path}...\n")
            if os.path.isfile(config_file):
                shutil.copy2(config_file, new_config_file_path)
                self.status_text.insert(tk.END, "Config file copied successfully.\n")
            else:
                self.status_text.insert(tk.END, f"Did not find a config file at {config_file}\n")

    def move_default_config(self):
        config_dir = self.config_dir_entry.get()
        if not config_dir:
            messagebox.showerror("Error", "Please select the config directory.")
            return

        # Save paths and checkbox states to config file
        self.save_config(config_dir, self.shortcut_dir_entry.get(), max(10, int(self.delay_entry.get())), self.dark_mode_var.get(), {k: v.get() for k, v in self.bard_checkbuttons.items()}, self.lightamp_check_var.get(), self.lightamp_entry.get())

        config_file = os.path.join(os.path.expanduser('~'), "Documents", "My Games", "FINAL FANTASY XIV - A Realm Reborn", "FFXIV.cfg")
        default_config_file_path = os.path.join(config_dir, "default.cfg")

        self.status_text.insert(tk.END, "Moving default config file...\n")

        if os.path.isfile(default_config_file_path):
            shutil.copy2(default_config_file_path, config_file)
            self.status_text.insert(tk.END, "Default config file moved successfully.\n")
        else:
            self.status_text.insert(tk.END, f"Did not find a default config file at {default_config_file_path}\n")

    def save_config(self, config_dir, shortcut_dir, delay, dark_mode, bard_checkbuttons, lightamp_check, lightamp_location):
        config_data = {
            "config_dir": config_dir,
            "shortcut_dir": shortcut_dir,
            "delay": delay,
            "dark_mode": dark_mode,
            "bard_checkbuttons": bard_checkbuttons,
            "lightamp_check": lightamp_check,
            "lightamp_location": lightamp_location
        }
        with open(CONFIG_FILE, 'w') as f:
            json.dump(config_data, f)

    def load_config(self):
        if os.path.isfile(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    return json.load(f)
            except json.JSONDecodeError:
                return {}
        return {}

    def clear_status(self):
        self.status_text.delete(1.0, tk.END)

    def save_settings(self):
        self.save_config(
            self.config_dir_entry.get(),
            self.shortcut_dir_entry.get(),
            max(10, int(self.delay_entry.get())),
            self.dark_mode_var.get(),
            {k: v.get() for k, v in self.bard_checkbuttons.items()},
            self.lightamp_check_var.get(),
            self.lightamp_entry.get()
        )
        self.status_text.insert(tk.END, "Settings saved.\n")

    def load_settings(self):
        self.config_data = self.load_config()
        if self.config_data:
            self.config_dir_entry.delete(0, tk.END)
            self.config_dir_entry.insert(0, self.config_data.get('config_dir', ''))
            self.shortcut_dir_entry.delete(0, tk.END)
            self.shortcut_dir_entry.insert(0, self.config_data.get('shortcut_dir', ''))
            self.delay_entry.delete(0, tk.END)
            self.delay_entry.insert(0, str(self.config_data.get('delay', 10)))
            self.dark_mode_var.set(self.config_data.get('dark_mode', False))
            self.toggle_dark_mode()  # Set theme based on loaded config
            if 'bard_checkbuttons' in self.config_data:
                self.populate_shortcuts(self.config_data['bard_checkbuttons'])
            else:
                self.populate_shortcuts()
            self.lightamp_check_var.set(self.config_data.get('lightamp_check', False))
            self.lightamp_entry.insert(0, self.config_data.get('lightamp_location', ''))
        self.status_text.insert(tk.END, "Settings loaded.\n")

    def reset_configuration(self):
        if messagebox.askokcancel("Reset Configuration", "Are you sure you want to reset the configuration to default settings?"):
            self.config_dir_entry.delete(0, tk.END)
            self.shortcut_dir_entry.delete(0, tk.END)
            self.delay_entry.delete(0, tk.END)
            self.delay_entry.insert(0, "10")
            self.dark_mode_var.set(False)
            self.toggle_dark_mode()  # Reset to light theme
            for widget in self.bard_buttons_frame.winfo_children():
                widget.destroy()
            self.bard_checkbuttons = {}
            self.lightamp_check_var.set(False)
            self.lightamp_entry.delete(0, tk.END)
            self.status_text.insert(tk.END, "Configuration reset to default.\n")

    def show_context_menu(self, event, bard_name):
        context_menu = Menu(self.root, tearoff=0)
        context_menu.add_command(label="Launch", command=lambda: self.launch_bard(self.shortcut_dir_entry.get(), bard_name))
        context_menu.add_command(label="Copy Config", command=lambda: self.copy_config(bard_name))
        context_menu.add_command(label="Change Icon", command=lambda: self.change_icon(bard_name))
        context_menu.add_command(label="Rename", command=lambda: self.rename_shortcut(bard_name))
        context_menu.tk_popup(event.x_root, event.y_root)

    def change_icon(self, bard_name):
        file_path = filedialog.askopenfilename(filetypes=[("PNG files", "*.png"), ("All files", "*.*")])
        if file_path:
            shortcut_path = os.path.join(self.shortcut_dir_entry.get(), f"{bard_name}.lnk")
            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.IconLocation = file_path
            shortcut.save()
            self.populate_shortcuts()

    def rename_shortcut(self, bard_name):
        new_name = simpledialog.askstring("Rename Shortcut", f"Enter new name for {bard_name}:")
        if new_name:
            shortcut_dir = self.shortcut_dir_entry.get()
            old_shortcut_path = os.path.join(shortcut_dir, f"{bard_name}.lnk")
            new_shortcut_path = os.path.join(shortcut_dir, f"{new_name}.lnk")
            if os.path.exists(new_shortcut_path):
                messagebox.showerror("Error", f"A shortcut with the name {new_name} already exists.")
                return
            os.rename(old_shortcut_path, new_shortcut_path)

            config_dir = self.config_dir_entry.get()
            old_config_path = os.path.join(config_dir, f"{bard_name}.cfg")
            new_config_path = os.path.join(config_dir, f"{new_name}.cfg")
            if os.path.exists(old_config_path):
                os.rename(old_config_path, new_config_path)

            self.populate_shortcuts()

if __name__ == "__main__":
    root = ThemedTk(theme="clam")
    app = BardLauncherGUI(root)
    root.mainloop()
