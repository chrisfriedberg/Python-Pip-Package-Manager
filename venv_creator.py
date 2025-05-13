import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, Listbox
import os
import subprocess
import sys
import venv
import json
import importlib.util
import sysconfig
import threading
import datetime
import webbrowser
import platform

# Store venv history in the script root folder
VENVS_HISTORY_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "venvs_history.json")

class VenvCreatorDialog(ctk.CTkToplevel):
    def __init__(self, master, icon_path=None):
        super().__init__(master)
        self.title("Virtual Environment Creator")
        self.icon_path = icon_path
        self.transient(master)
        self.grab_set()
        self.geometry("600x500")
        
        # Set icon if available
        if self.icon_path and os.path.exists(self.icon_path):
            try:
                self.iconbitmap(self.icon_path)
            except Exception:
                pass

        # Location Frame
        location_frame = ctk.CTkFrame(self)
        location_frame.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(location_frame, text="Virtual Env (venv) Destination:").pack(side="left", padx=5)
        self.location_entry = ctk.CTkEntry(location_frame, width=300)
        self.location_entry.pack(side="left", padx=5, fill="x", expand=True)
        browse_btn = ctk.CTkButton(location_frame, text="Browse", width=80, command=self.browse_location)
        browse_btn.pack(side="right", padx=5)

        # Requirements Frame
        req_frame = ctk.CTkFrame(self)
        req_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Requirements Box Frame (with border and header)
        req_box_frame = ctk.CTkFrame(req_frame, fg_color="#23272E", border_color="#1A73E8", border_width=2, corner_radius=8)
        req_box_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        req_label = ctk.CTkLabel(req_box_frame, text="Requirements List", font=("Arial", 13, "bold"), text_color="#1A73E8")
        req_label.pack(anchor="nw", padx=10, pady=(8, 2))
        
        # Listbox and Buttons Frame
        list_btn_frame = ctk.CTkFrame(req_box_frame, fg_color="transparent")
        list_btn_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Listbox
        self.req_listbox = Listbox(list_btn_frame, height=7, width=42, bg="#181A20", fg="white", 
                                 borderwidth=0, highlightthickness=0, exportselection=False, font=("Consolas", 12))
        self.req_listbox.pack(side="left", fill="both", expand=True, padx=(0, 10), pady=5)
        
        # Buttons Frame (fixed width, visible background)
        btn_frame = ctk.CTkFrame(list_btn_frame, fg_color="#23272E", width=180)
        btn_frame.pack(side="right", fill="y", padx=5, pady=5)
        btn_frame.pack_propagate(False)
        
        btn_width = 160
        btn_height = 36
        add_btn = ctk.CTkButton(btn_frame, text="Add", width=btn_width, height=btn_height, fg_color="#27ae60", hover_color="#219150", command=self.add_requirement)
        add_btn.pack(pady=(0, 8))
        remove_btn = ctk.CTkButton(btn_frame, text="Remove", width=btn_width, height=btn_height, fg_color="#e74c3c", hover_color="#b93a2b", command=self.remove_requirement)
        remove_btn.pack(pady=(0, 8))
        import_btn = ctk.CTkButton(btn_frame, text="Import requirements", width=btn_width, height=btn_height, fg_color="#1565c0", hover_color="#0d47a1", command=self.import_requirements)
        import_btn.pack(pady=(0, 8))
        scrape_btn = ctk.CTkButton(btn_frame, text="Scrape for requirements", width=btn_width, height=btn_height, fg_color="#1565c0", hover_color="#0d47a1", command=self.scrape_for_requirements)
        scrape_btn.pack(pady=(0, 8))
        
        # Tooltips
        try:
            from CTkToolTip import CTkToolTip
            CTkToolTip(add_btn, message="Add a package to the requirements list")
            CTkToolTip(remove_btn, message="Remove the selected package from the list")
            CTkToolTip(import_btn, message="Import packages from a requirements.txt file")
            CTkToolTip(scrape_btn, message="Scan your project for imports and auto-populate requirements")
        except Exception:
            pass

        # Add Export requirements button below the other buttons in btn_frame
        export_btn = ctk.CTkButton(btn_frame, text="Export requirements", width=btn_width, height=btn_height, fg_color="#1565c0", hover_color="#0d47a1", command=self.export_requirements)
        export_btn.pack(side="bottom", pady=(16, 0), fill="x")

        # Bottom Buttons Frame
        bottom_frame = ctk.CTkFrame(self)
        bottom_frame.pack(fill="x", padx=10, pady=10)
        
        associate_btn = ctk.CTkButton(bottom_frame, text="Associate PyScript", width=150, 
                                    command=self.associate_pyscript)
        associate_btn.pack(side="left", padx=5)
        
        venv_history_btn = ctk.CTkButton(bottom_frame, text="Venv History", width=150, fg_color="#1565c0", hover_color="#0d47a1", command=self.open_venv_history)
        venv_history_btn.pack(side="left", padx=5)
        
        # Button alignment frame for right side
        right_btn_frame = ctk.CTkFrame(bottom_frame, fg_color="transparent")
        right_btn_frame.pack(side="right", padx=0, pady=0)
        
        create_btn = ctk.CTkButton(right_btn_frame, text="Create venv", width=150, 
                                 command=self.create_venv, fg_color="#1A73E8", hover_color="#1557B0")
        create_btn.pack(side="top", padx=5, pady=(0, 8), anchor="e")
        close_btn = ctk.CTkButton(right_btn_frame, text="Close", width=150, fg_color="#e74c3c", hover_color="#b93a2b", command=self.destroy)
        close_btn.pack(side="top", padx=5, pady=(0, 0), anchor="e")
        
        # Increase window height to fit new button
        self.geometry("600x560")

        # Spinner/progress indicator
        self.progress_label = ctk.CTkLabel(self, text="", font=("Arial", 12, "italic"), text_color="#1A73E8")
        self.progress_label.pack_forget()
        self.progress_bar = ctk.CTkProgressBar(self, orientation="horizontal", mode="indeterminate", width=300)
        self.progress_bar.pack_forget()

    def browse_location(self):
        location = filedialog.askdirectory(title="Select Virtual Environment Location")
        if location:
            self.location_entry.delete(0, tk.END)
            self.location_entry.insert(0, location)

    def add_requirement(self):
        dialog = ctk.CTkInputDialog(text="Enter package name:", title="Add Requirement")
        package = dialog.get_input()
        if package and package.strip():
            pkg = package.strip()
            if is_stdlib_module(pkg):
                self.show_error(f"'{pkg}' is a standard library module and does not need to be installed via pip.")
                return
            self.req_listbox.insert(tk.END, pkg)

    def remove_requirement(self):
        selection = self.req_listbox.curselection()
        if selection:
            self.req_listbox.delete(selection)

    def associate_pyscript(self):
        # This is a placeholder for the PyScript association functionality
        # You can implement the actual PyScript association logic here
        pass

    def show_spinner(self, message="Working..."):
        self.progress_label.configure(text=message)
        self.progress_label.pack(pady=(0, 5))
        self.progress_bar.pack(pady=(0, 10))
        self.progress_bar.start()
        self.update()

    def hide_spinner(self):
        self.progress_label.pack_forget()
        self.progress_bar.stop()
        self.progress_bar.pack_forget()
        self.update()

    def create_venv(self):
        parent_location = self.location_entry.get().strip()
        if not parent_location:
            self.show_error("Please select a location for the virtual environment.")
            return

        venv_path = os.path.join(parent_location, "venv")
        if os.path.exists(venv_path):
            # Prompt for overwrite
            confirm = ctk.CTkToplevel(self)
            confirm.title("Overwrite venv?")
            confirm.geometry("400x180")
            confirm.transient(self)
            confirm.grab_set()
            ctk.CTkLabel(confirm, text=f"A 'venv' folder already exists in this location. Overwrite?", wraplength=350).pack(pady=20, padx=20)
            btn_frame = ctk.CTkFrame(confirm, fg_color="transparent")
            btn_frame.pack(pady=10)
            result = {'ok': False}
            def on_ok():
                result['ok'] = True
                confirm.destroy()
            def on_cancel():
                confirm.destroy()
            ctk.CTkButton(btn_frame, text="Overwrite", command=on_ok).pack(side="left", padx=10)
            ctk.CTkButton(btn_frame, text="Cancel", command=on_cancel).pack(side="right", padx=10)
            confirm.wait_window()
            if not result['ok']:
                return
            # Remove existing venv folder
            import shutil
            try:
                shutil.rmtree(venv_path)
            except Exception as e:
                self.show_error(f"Could not remove existing venv: {e}")
                return

        def venv_task():
            self.show_spinner("Creating virtual environment...")
            try:
                # Create the virtual environment in the 'venv' subfolder
                import venv
                venv.create(venv_path, with_pip=True)
                # Get requirements from listbox, filter stdlib
                requirements = [pkg for pkg in self.req_listbox.get(0, tk.END) if not is_stdlib_module(pkg)]
                skipped = [pkg for pkg in self.req_listbox.get(0, tk.END) if is_stdlib_module(pkg)]
                if requirements:
                    # Create requirements.txt
                    req_file = os.path.join(venv_path, "requirements.txt")
                    with open(req_file, "w") as f:
                        f.write("\n".join(requirements))
                    # Install requirements and capture output
                    pip_path = os.path.join(venv_path, "Scripts" if sys.platform == "win32" else "bin", "pip")
                    if sys.platform == "win32":
                        pip_path += ".exe"
                    success = []
                    failed = []
                    for pkg in requirements:
                        try:
                            result = subprocess.run([pip_path, "install", pkg], capture_output=True, text=True, check=True)
                            success.append(pkg)
                        except subprocess.CalledProcessError as e:
                            failed.append(pkg)
                    self.hide_spinner()
                    self.log_venv_creation(venv_path)
                    self.show_install_summary(success, failed, skipped)
                    # Clear UI after success
                    self.req_listbox.delete(0, tk.END)
                    self.location_entry.delete(0, tk.END)
                else:
                    self.hide_spinner()
                    self.log_venv_creation(venv_path)
                    self.show_success("Virtual Environment Created!\n(No installable packages specified.)")
            except Exception as e:
                self.hide_spinner()
                self.show_error(f"Error creating virtual environment: {str(e)}")
        threading.Thread(target=venv_task, daemon=True).start()

    def show_error(self, message):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Error")
        dialog.geometry("400x150")
        dialog.transient(self)
        dialog.grab_set()
        dialog.lift()
        dialog.focus_force()
        self.center_dialog(dialog)
        ctk.CTkLabel(dialog, text=message, wraplength=350).pack(pady=20, padx=20)
        ctk.CTkButton(dialog, text="OK", command=dialog.destroy).pack(pady=10)

    def show_success(self, message):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Success")
        dialog.geometry("400x150")
        dialog.transient(self)
        dialog.grab_set()
        dialog.lift()
        dialog.focus_force()
        self.center_dialog(dialog)
        ctk.CTkLabel(dialog, text=message, wraplength=350).pack(pady=20, padx=20)
        ctk.CTkButton(dialog, text="OK", command=dialog.destroy).pack(pady=10)

    def show_install_summary(self, success, failed, skipped):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Venv Creation Summary")
        dialog.geometry("450x350")
        dialog.transient(self)
        dialog.grab_set()
        msg = ""
        if success:
            msg += "✔️ Packages successfully installed:\n" + "\n".join(success) + "\n\n"
        if failed:
            msg += "⛔ Packages failed to install:\n" + "\n".join(failed) + "\n\n"
        if skipped:
            msg += "ℹ️ Skipped standard library modules:\n" + "\n".join(skipped) + "\n\n"
        if not msg:
            msg = "No packages were specified."
        ctk.CTkLabel(dialog, text=msg, wraplength=400, justify="left").pack(pady=20, padx=20)
        ctk.CTkButton(dialog, text="OK", command=dialog.destroy).pack(pady=10)

    def import_requirements(self):
        file_path = filedialog.askopenfilename(title="Select requirements.txt", filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if not file_path:
            return
        skipped = []
        added = 0
        with open(file_path, "r") as f:
            for line in f:
                pkg = line.strip()
                if not pkg or pkg.startswith("#"):
                    continue
                if is_stdlib_module(pkg):
                    skipped.append(pkg)
                    continue
                # Avoid duplicates
                current = set(self.req_listbox.get(0, tk.END))
                if pkg in current:
                    continue
                self.req_listbox.insert(tk.END, pkg)
                added += 1
        msg = f"Added {added} packages from requirements.txt."
        if skipped:
            msg += "\nSkipped standard library modules: " + ", ".join(skipped)
        self.show_success(msg)

    def scrape_for_requirements(self):
        # Let user multi-select .py and .pyw files to analyze
        from tkinter import filedialog
        import subprocess
        import sys
        file_paths = filedialog.askopenfilenames(title="Select Python files to analyze", filetypes=[("Python files", "*.py;*.pyw")])
        if not file_paths:
            return
        script_path = os.path.join(os.path.dirname(__file__), "requirements_generator.py")
        if not os.path.exists(script_path):
            self.show_error("requirements_generator.py not found in the app directory.")
            return
        # Use the directory of the first selected file
        target_dir = os.path.dirname(file_paths[0])
        def scrape_task():
            self.show_spinner("Scraping requirements from selected files...")
            try:
                # Pass the directory to requirements_generator.py
                result = subprocess.run([sys.executable, script_path, target_dir], capture_output=True, text=True)
                # Read the generated requirements.txt in the same directory as the first file
                req_path = os.path.join(target_dir, "requirements.txt")
                if os.path.exists(req_path):
                    with open(req_path, "r", encoding="utf-8") as f:
                        pkgs = [line.strip() for line in f if line.strip() and not line.startswith('#')]
                    # Clear and repopulate the listbox
                    self.req_listbox.delete(0, tk.END)
                    skipped = []
                    for pkg in pkgs:
                        if is_stdlib_module(pkg):
                            skipped.append(pkg)
                            continue
                        self.req_listbox.insert(tk.END, pkg)
                    msg = f"Scraped and added {len(pkgs) - len(skipped)} packages."
                    if skipped:
                        msg += "\nSkipped standard library modules: " + ", ".join(skipped)
                    self.hide_spinner()
                    self.show_success(msg)
                else:
                    self.hide_spinner()
                    self.show_error("requirements.txt was not generated.")
            except Exception as e:
                self.hide_spinner()
                self.show_error(f"Error scraping requirements: {e}")
        threading.Thread(target=scrape_task, daemon=True).start()

    def log_venv_creation(self, venv_path):
        parent_dir = os.path.dirname(venv_path)
        entry = {
            "date": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "parent": os.path.basename(parent_dir),
            "venv_path": venv_path
        }
        try:
            if os.path.exists(VENVS_HISTORY_FILE):
                with open(VENVS_HISTORY_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
            else:
                data = []
            data.append(entry)
            with open(VENVS_HISTORY_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            print(f"Error logging venv creation: {e}")

    def open_venv_history(self):
        # Dialog for venv history
        dialog = ctk.CTkToplevel(self)
        dialog.title("Venv Creation History")
        dialog.geometry("820x480")
        dialog.transient(self)
        dialog.grab_set()
        dialog.lift()
        dialog.focus_force()
        self.center_dialog(dialog)
        
        # Hide deleted venvs checkbox (above table)
        hide_deleted_var = tk.BooleanVar(value=False)
        hide_cb = ctk.CTkCheckBox(dialog, text="Hide deleted venvs", variable=hide_deleted_var)
        hide_cb.pack(anchor="w", padx=20, pady=(10, 0))

        # Table frame
        table_frame = ctk.CTkFrame(dialog, fg_color="#23272E")
        table_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Header row
        headers = [
            ("Date Added", 140),
            ("Date Removed", 140),
            ("Script Root", 140),
            ("Venv Path", 280),
            ("", 80)  # For delete button
        ]
        for col, (header, width) in enumerate(headers):
            ctk.CTkLabel(table_frame, text=header, font=("Arial", 12, "bold"), width=width, anchor="w").grid(row=0, column=col, padx=5, pady=(0, 4), sticky="w")

        def refresh_history():
            # Remove all rows except header
            for widget in table_frame.winfo_children():
                info = widget.grid_info()
                if info['row'] != 0:
                    widget.destroy()
            # Load data
            try:
                if os.path.exists(VENVS_HISTORY_FILE):
                    with open(VENVS_HISTORY_FILE, "r", encoding="utf-8") as f:
                        data = json.load(f)
                else:
                    data = []
            except Exception as e:
                data = []
            if hide_deleted_var.get():
                data = [entry for entry in data if not entry.get("date_deleted")]
            for i, entry in enumerate(data):
                row = i + 1  # header is row 0
                date_added = entry.get("date", "")
                date_removed = entry.get("date_deleted", "")
                script_root = entry.get("parent", "")
                venv_path = entry.get("venv_path", "")
                ctk.CTkLabel(table_frame, text=date_added, width=140, anchor="w").grid(row=row, column=0, padx=5, sticky="w")
                ctk.CTkLabel(table_frame, text=date_removed, width=140, anchor="w").grid(row=row, column=1, padx=5, sticky="w")
                ctk.CTkLabel(table_frame, text=script_root, width=140, anchor="w").grid(row=row, column=2, padx=5, sticky="w")
                def open_folder(path):
                    folder = os.path.dirname(path)
                    if platform.system() == "Windows":
                        os.startfile(folder)
                    elif platform.system() == "Darwin":
                        subprocess.Popen(["open", folder])
                    else:
                        subprocess.Popen(["xdg-open", folder])
                def make_link_callback(p=venv_path):
                    return lambda e: open_folder(p)
                venv_path_label = ctk.CTkLabel(table_frame, text=venv_path, width=280, anchor="w", text_color="#1A73E8", cursor="hand2")
                venv_path_label.grid(row=row, column=3, padx=5, sticky="w")
                venv_path_label.bind("<Button-1>", make_link_callback())
                if not date_removed:
                    del_btn = ctk.CTkButton(table_frame, text="Delete", width=70, fg_color="#e74c3c", hover_color="#b93a2b", command=lambda idx=i: self.delete_venv_history(idx, dialog, hide_deleted_var))
                    del_btn.grid(row=row, column=4, padx=5, sticky="w")
        hide_cb.configure(command=refresh_history)
        refresh_history()

        # Button dimensions
        btn_width = 180
        btn_height = 32
        # Close button (bottom right, under browse_remove_btn)
        close_btn = ctk.CTkButton(dialog, text="Close", width=btn_width, height=btn_height, fg_color="#e74c3c", hover_color="#b93a2b", command=dialog.destroy)
        close_btn.place(relx=1.0, rely=1.0, x=-20, y=-20-btn_height-10, anchor="se")
        # Browse To Remove venv button (bottom right, blue)
        browse_remove_btn = ctk.CTkButton(dialog, text="Browse To Remove venv", width=btn_width, height=btn_height, fg_color="#1565c0", hover_color="#0d47a1", command=lambda: self.browse_remove_venv(dialog, refresh_history))
        browse_remove_btn.place(relx=1.0, rely=1.0, x=-20, y=-20, anchor="se")

    def delete_venv_history(self, idx, dialog, hide_deleted_var):
        try:
            if os.path.exists(VENVS_HISTORY_FILE):
                with open(VENVS_HISTORY_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
            else:
                data = []
            if 0 <= idx < len(data):
                data[idx]["date_deleted"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                with open(VENVS_HISTORY_FILE, "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=2)
            dialog.destroy()
            self.open_venv_history()
        except Exception as e:
            self.show_error(f"Error deleting venv history entry: {e}")

    def browse_remove_venv(self, dialog, refresh_history):
        from tkinter import filedialog
        import shutil
        venv_dir = filedialog.askdirectory(title="Select venv folder to delete")
        if not venv_dir:
            return
        # Confirm deletion only if the directory is named 'venv' or '.venv'
        if not os.path.basename(venv_dir) in ['venv', '.venv']:
            self.show_error("You can only delete directories named 'venv' or '.venv'.")
            return
        # Confirm deletion
        confirm = ctk.CTkToplevel(self)
        confirm.title("Confirm Delete venv")
        confirm.geometry("420x180")
        confirm.transient(dialog)
        confirm.grab_set()
        confirm.lift()
        confirm.focus_force()
        self.center_dialog(confirm, parent=dialog)
        ctk.CTkLabel(confirm, text=f"Are you sure you want to delete this venv?\n{venv_dir}", wraplength=380).pack(pady=20, padx=20)
        btn_frame = ctk.CTkFrame(confirm, fg_color="transparent")
        btn_frame.pack(pady=10)
        result = {'ok': False}
        def on_ok():
            result['ok'] = True
            confirm.destroy()
        def on_cancel():
            confirm.destroy()
        ctk.CTkButton(btn_frame, text="Delete", fg_color="#e74c3c", hover_color="#b93a2b", command=on_ok).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="Cancel", command=on_cancel).pack(side="right", padx=10)
        confirm.wait_window()
        if not result['ok']:
            return
        # Delete the folder
        try:
            shutil.rmtree(venv_dir)
            self.log_venv_deletion(venv_dir)
            self.show_success(f"Deleted venv: {venv_dir}")
            refresh_history()
        except Exception as e:
            self.show_error(f"Error deleting venv: {e}")

    def log_venv_deletion(self, venv_path):
        try:
            if os.path.exists(VENVS_HISTORY_FILE):
                with open(VENVS_HISTORY_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
            else:
                data = []
            found = False
            for entry in data:
                if entry["venv_path"] == venv_path:
                    entry["date_deleted"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    found = True
            if not found:
                # If not found, add a new entry for deletion
                parent_dir = os.path.dirname(venv_path)
                entry = {
                    "date": "",
                    "parent": os.path.basename(parent_dir),
                    "venv_path": venv_path,
                    "date_deleted": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
                data.append(entry)
            with open(VENVS_HISTORY_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            print(f"Error logging venv deletion: {e}")

    def center_dialog(self, dialog, parent=None):
        dialog.update_idletasks()
        if parent is None:
            parent = self
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_w = parent.winfo_width()
        parent_h = parent.winfo_height()
        win_w = dialog.winfo_width()
        win_h = dialog.winfo_height()
        x = parent_x + (parent_w // 2) - (win_w // 2)
        y = parent_y + (parent_h // 2) - (win_h // 2)
        dialog.geometry(f"+{x}+{y}")

    def export_requirements(self):
        file_path = filedialog.asksaveasfilename(title="Export requirements.txt", defaultextension=".txt", filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if not file_path:
            return
        pkgs = [self.req_listbox.get(idx) for idx in range(self.req_listbox.size())]
        with open(file_path, "w") as f:
            f.write("\n".join(pkgs))
        self.show_success(f"Exported requirements to {file_path}")

# Helper to check if a module is part of the standard library
def is_stdlib_module(module_name):
    if module_name.lower() == "pillow":  # Pillow is not stdlib
        return False
    try:
        # Try to find the module spec
        spec = importlib.util.find_spec(module_name)
        if spec is None:
            return False  # Not found at all
        # Check if it's in the stdlib path
        stdlib_path = sysconfig.get_paths()["stdlib"]
        if spec.origin and spec.origin.startswith(stdlib_path):
            return True
    except Exception:
        pass
    return False 