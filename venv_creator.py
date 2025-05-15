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

# Store environment history in the script root folder
ENV_HISTORY_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "venvs_history.json")

class VenvCreatorDialog(ctk.CTkToplevel):
    def __init__(self, master, icon_path=None):
        super().__init__(master)
        self.title("Virtual Environment Creator")
        self.icon_path = icon_path
        self.transient(master)
        
        # Set size but don't show window yet
        self.geometry("600x500")
        self.withdraw()  # Hide window initially
        
        # Initialize selected_venv_path to None
        self.selected_venv_path = None
        
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
        self.location_entry.bind("<KeyRelease>", self.check_create_button_state)
        
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
        
        # Add "Add Packages to Established Venv" button
        add_to_existing_btn = ctk.CTkButton(btn_frame, text="Add to Existing Venv", width=btn_width, height=btn_height, fg_color="#8E44AD", hover_color="#6C3483", command=self.add_packages_to_existing_venv)
        add_to_existing_btn.pack(pady=(0, 8))
        
        # Tooltips
        try:
            from CTkToolTip import CTkToolTip
            CTkToolTip(add_btn, message="Add a package to the requirements list")
            CTkToolTip(remove_btn, message="Remove the selected package from the list")
            CTkToolTip(import_btn, message="Import packages from a requirements.txt file")
            CTkToolTip(scrape_btn, message="Scan your project for imports and auto-populate requirements")
            CTkToolTip(add_to_existing_btn, message="Add packages to an existing virtual environment")
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
        
        # Create venv button (initially disabled)
        self.create_btn = ctk.CTkButton(right_btn_frame, text="Create venv", width=150, 
                                 command=self.create_venv, fg_color="#1A73E8", hover_color="#1557B0", 
                                 state="disabled")  # Start disabled until location is entered
        self.create_btn.pack(side="top", padx=5, pady=(0, 8), anchor="e")
        
        close_btn = ctk.CTkButton(right_btn_frame, text="Close", width=150, fg_color="#e74c3c", hover_color="#b93a2b", command=self.destroy)
        close_btn.pack(side="top", padx=5, pady=(0, 0), anchor="e")
        
        # Increase window height to fit new button
        self.geometry("600x560")

        # Spinner/progress indicator
        self.progress_label = ctk.CTkLabel(self, text="", font=("Arial", 12, "italic"), text_color="#1A73E8")
        self.progress_label.pack_forget()
        self.progress_bar = ctk.CTkProgressBar(self, orientation="horizontal", mode="indeterminate", width=300)
        self.progress_bar.pack_forget()
        
        # Center window before showing it
        self.center_on_screen()
        
        # Now grab focus and make visible
        self.grab_set()
        self.deiconify()  # Show window after it's fully configured and centered
    
    def check_create_button_state(self, event=None):
        """Enable or disable the Create venv button based on whether a location is entered"""
        location = self.location_entry.get().strip()
        if location:
            self.create_btn.configure(state="normal")
        else:
            self.create_btn.configure(state="disabled")
    
    def center_on_screen(self):
        """Center this dialog on the screen"""
        # Make sure all widgets are properly realized
        self.update_idletasks()
        
        # Get screen dimensions
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        # Get window size
        width = self.winfo_width()
        height = self.winfo_height()
        
        # Calculate position
        x = max(0, int((screen_width - width) / 2))
        y = max(0, int((screen_height - height) / 2))
        
        # Position window
        self.geometry(f"+{x}+{y}")
        
        # Force update to ensure window appears at the correct position immediately
        self.update()
            
    def browse_location(self):
        location = filedialog.askdirectory(title="Select Virtual Environment Location")
        if location:
            self.location_entry.delete(0, tk.END)
            self.location_entry.insert(0, location)
            self.check_create_button_state()  # Update button state after browsing

    def add_requirement(self):
        dialog = ctk.CTkInputDialog(text="Enter package name:", title="Add Requirement")
        package = dialog.get_input()
        if package and package.strip():
            pkg = package.strip()
            if pkg.lower() == "pyside6":
                # Allow PySide6 to be added to the list for user visibility, but inform them it's already included
                self.show_info("Note: PySide6 is automatically installed in all virtual environments created by this tool, but it's been added to your list for visibility.")
                self.req_listbox.insert(tk.END, pkg)
                return
            if is_stdlib_module(pkg):
                self.show_error(f"'{pkg}' is a standard library module and does not need to be installed via pip.")
                return
            self.req_listbox.insert(tk.END, pkg)

    def remove_requirement(self):
        selection = self.req_listbox.curselection()
        if selection:
            self.req_listbox.delete(selection)

    def associate_pyscript(self):
        """Associate a Python script with a virtual environment and create launcher files"""
        # Dialog for selecting script and venv
        dialog = ctk.CTkToplevel(self)
        dialog.title("Associate Script with Virtual Environment")
        
        # Configure but don't show yet
        dialog.geometry("600x480")
        dialog.transient(self)
        dialog.withdraw()  # Hide initially
        
        # Set icon if available
        if self.icon_path and os.path.exists(self.icon_path):
            try:
                dialog.iconbitmap(self.icon_path)
            except Exception:
                pass
        
        # Script selection frame
        script_frame = ctk.CTkFrame(dialog)
        script_frame.pack(fill="x", padx=20, pady=(20, 10))
        
        ctk.CTkLabel(script_frame, text="Select Python Script:").pack(side="left", padx=5)
        script_entry = ctk.CTkEntry(script_frame, width=300)
        script_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        def browse_script():
            file_path = filedialog.askopenfilename(
                title="Select Python Script", 
                filetypes=[("Python files", "*.py;*.pyw"), ("All files", "*.*")]
            )
            if file_path:
                script_entry.delete(0, tk.END)
                script_entry.insert(0, file_path)
                
                # Auto-fill the parent directory as default venv location
                parent_dir = os.path.dirname(file_path)
                venv_entry.delete(0, tk.END)
                venv_entry.insert(0, os.path.join(parent_dir, "venv"))
        
        browse_script_btn = ctk.CTkButton(script_frame, text="Browse", width=80, command=browse_script)
        browse_script_btn.pack(side="right", padx=5)
        
        # Venv selection frame
        venv_frame = ctk.CTkFrame(dialog)
        venv_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(venv_frame, text="Select Virtual Environment:").pack(side="left", padx=5)
        venv_entry = ctk.CTkEntry(venv_frame, width=300)
        venv_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        def browse_venv():
            venv_dir = filedialog.askdirectory(title="Select Virtual Environment Directory")
            if venv_dir:
                venv_entry.delete(0, tk.END)
                venv_entry.insert(0, venv_dir)
        
        browse_venv_btn = ctk.CTkButton(venv_frame, text="Browse", width=80, command=browse_venv)
        browse_venv_btn.pack(side="right", padx=5)
        
        # Options frame
        options_frame = ctk.CTkFrame(dialog)
        options_frame.pack(fill="x", padx=20, pady=10)
        
        # System tray option
        systray_var = tk.BooleanVar(value=True)
        systray_cb = ctk.CTkCheckBox(options_frame, text="Add system tray icon with exit option", variable=systray_var)
        systray_cb.pack(anchor="w", padx=5, pady=5)
        
        # Add startup with Windows option
        startup_var = tk.BooleanVar(value=False)
        startup_cb = ctk.CTkCheckBox(options_frame, text="Auto-start application at Windows login", variable=startup_var)
        startup_cb.pack(anchor="w", padx=5, pady=5)
        
        # Information text
        info_frame = ctk.CTkFrame(dialog)
        info_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        info_text = (
            "This will create the following files in the script's directory:\n\n"
            "1. .venv-association - Associates the script with the venv\n"
            "2. run_this.vbs - Silent launcher with no terminal window\n"
            "3. run_this.bat - Command line launcher for manual use\n"
            "4. launcher.py - System tray launcher with exit functionality\n\n"
            "You can create a shortcut to launcher.py and customize its icon."
        )
        
        ctk.CTkLabel(info_frame, text=info_text, wraplength=550, justify="left").pack(pady=20, padx=20, fill="both", expand=True)
        
        # Buttons frame
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=(0, 20))
        
        # Add the delete association function
        def delete_association_browse():
            # Close the current dialog
            if dialog.winfo_exists():
                dialog.destroy()
                
            # Open a new dialog for deletion
            del_dialog = ctk.CTkToplevel(self)
            del_dialog.title("Delete Association Files")
            
            # Configure but don't show yet
            del_dialog.geometry("500x300")
            del_dialog.transient(self)
            del_dialog.withdraw()  # Hide initially
            
            # Set icon if available
            if self.icon_path and os.path.exists(self.icon_path):
                try:
                    del_dialog.iconbitmap(self.icon_path)
                except Exception:
                    pass
            
            # Label and explanation
            ctk.CTkLabel(
                del_dialog, 
                text="Select the directory containing association files to delete",
                font=("Arial", 12, "bold")
            ).pack(pady=(20, 5), padx=20)
            
            help_text = (
                "This will delete the following files if they exist in the selected directory:\n"
                "• .venv-association\n"
                "• run_this.vbs\n"
                "• run_this.bat\n"
                "• launcher.py\n"
                "• launcher_icon.ico"
            )
            
            ctk.CTkLabel(
                del_dialog, 
                text=help_text,
                justify="left",
                wraplength=450
            ).pack(pady=(0, 20), padx=20)
            
            # Directory selection frame
            dir_frame = ctk.CTkFrame(del_dialog)
            dir_frame.pack(fill="x", padx=20, pady=10)
            
            dir_entry = ctk.CTkEntry(dir_frame, width=350)
            dir_entry.pack(side="left", padx=5, fill="x", expand=True)
            
            def browse_dir():
                dir_path = filedialog.askdirectory(title="Select Directory with Association Files")
                if dir_path:
                    dir_entry.delete(0, tk.END)
                    dir_entry.insert(0, dir_path)
            
            browse_dir_btn = ctk.CTkButton(dir_frame, text="Browse", width=80, command=browse_dir)
            browse_dir_btn.pack(side="right", padx=5)
            
            # Button frame
            del_btn_frame = ctk.CTkFrame(del_dialog, fg_color="transparent")
            del_btn_frame.pack(fill="x", padx=20, pady=(20, 20))
            
            def perform_delete():
                dir_path = dir_entry.get().strip()
                if not dir_path:
                    self.show_error("Please select a directory first.")
                    return
                
                if not os.path.exists(dir_path) or not os.path.isdir(dir_path):
                    self.show_error(f"The selected path is not a valid directory:\n{dir_path}")
                    return
                
                result = self.delete_association_files(dir_path, show_results=True)
                
                if result["deleted"]:
                    message = "Successfully deleted the following files:\n• " + "\n• ".join(result["deleted"])
                    if result["missing"]:
                        message += "\n\nThe following files were not found:\n• " + "\n• ".join(result["missing"])
                    if result["failed"]:
                        message += "\n\nFailed to delete the following files:\n• " + "\n• ".join(result["failed"])
                    self.show_success(message)
                else:
                    if result["missing"]:
                        self.show_error(f"No association files were found in:\n{dir_path}")
                    elif result["failed"]:
                        self.show_error(f"Failed to delete association files:\n• " + "\n• ".join(result["failed"]))
                
                # Close the dialog
                if del_dialog.winfo_exists():
                    del_dialog.destroy()
            
            # Delete and Cancel buttons
            delete_btn = ctk.CTkButton(
                del_btn_frame,
                text="Delete Files",
                width=150,
                fg_color="#e74c3c",
                hover_color="#b93a2b",
                command=perform_delete
            )
            delete_btn.pack(side="left", padx=5)
            
            cancel_btn = ctk.CTkButton(
                del_btn_frame,
                text="Cancel",
                width=100,
                command=del_dialog.destroy
            )
            cancel_btn.pack(side="right", padx=5)
            
            # Center and show dialog
            self.center_dialog(del_dialog)
            del_dialog.grab_set()
            del_dialog.deiconify()
        
        def create_association():
            script_path = script_entry.get().strip()
            venv_path = venv_entry.get().strip()
            
            if not script_path or not os.path.exists(script_path):
                self.show_error("Please select a valid Python script.")
                return
            
            if not venv_path or not os.path.exists(venv_path):
                self.show_error("Please select a valid virtual environment directory.")
                return
                
            # Create the association using our new function
            result = self.create_association(
                script_path=script_path,
                venv_path=venv_path,
                add_systray=systray_var.get(),
                add_startup=startup_var.get()
            )
            
            if not result["success"]:
                self.show_error(result["error"])
                return
                
            # Get information from the result
            script_dir = result["script_dir"]
            files_created = result["files_created"]
            add_systray = result["add_systray"]
            add_startup = result["add_startup"]
                
            # Show success message with details
            success_msg = (
                f"Successfully created association files in:\n{script_dir}\n\n"
                "Files created:\n- " + "\n- ".join(files_created) + "\n\n"
            )
            
            if add_systray:
                success_msg += "You can now run your script by launching launcher.py\n"
                success_msg += "Right-click the tray icon to access the Exit option."
            else:
                success_msg += "You can now run your script by double-clicking run_this.vbs"
                
            if add_startup:
                success_msg += "\nThe application will also start automatically at Windows login."
                
            # Mention PySide6 installation
            if add_systray:
                success_msg += "\n\nNote: The system tray icon uses PySide6, which is automatically installed in all virtual environments created by this tool."
            
            # Show success dialog BEFORE destroying the association dialog
            success_dialog = ctk.CTkToplevel(self)
            success_dialog.title("Success")
            
            # Configure but don't show yet
            success_dialog.geometry("500x300")
            success_dialog.transient(self)
            success_dialog.withdraw()  # Hide initially
            
            # Set icon if available
            if self.icon_path and os.path.exists(self.icon_path):
                try:
                    success_dialog.iconbitmap(self.icon_path)
                except Exception:
                    pass
                    
            ctk.CTkLabel(success_dialog, text=success_msg, wraplength=480, justify="left").pack(pady=20, padx=20)
            
            def close_both_and_open_location():
                if success_dialog.winfo_exists():
                    success_dialog.destroy()
                if dialog.winfo_exists():
                    dialog.destroy()
                # Open the script directory
                try:
                    if os.path.exists(script_dir):
                        if platform.system() == "Windows":
                            os.startfile(script_dir)
                        elif platform.system() == "Darwin":
                            subprocess.Popen(["open", script_dir])
                        else:
                            subprocess.Popen(["xdg-open", script_dir])
                except Exception as e:
                    print(f"Error opening script directory: {e}")
                    
            ctk.CTkButton(success_dialog, text="OK", command=close_both_and_open_location).pack(pady=10)
        
            # Center and show dialog
            self.center_dialog(success_dialog)
            success_dialog.grab_set()
            success_dialog.deiconify()
        
        # Create Association button
        create_btn = ctk.CTkButton(
            btn_frame, 
            text="Create Association", 
            width=150, 
            fg_color="#1A73E8", 
            hover_color="#1557B0", 
            command=create_association
        )
        create_btn.pack(side="left", padx=5)
        
        # Delete Association button
        delete_btn = ctk.CTkButton(
            btn_frame, 
            text="Delete Association", 
            width=150, 
            fg_color="#e74c3c", 
            hover_color="#b93a2b", 
            command=delete_association_browse
        )
        delete_btn.pack(side="left", padx=5)
        
        # Cancel button
        cancel_btn = ctk.CTkButton(
            btn_frame, 
            text="Cancel", 
            width=100, 
            fg_color="#888888", 
            hover_color="#666666", 
            command=dialog.destroy
        )
        cancel_btn.pack(side="right", padx=5)
        
        # Center and show dialog
        self.center_dialog(dialog)
        dialog.grab_set()
        dialog.deiconify()

    def delete_association_files(self, script_dir, show_results=False):
        """Delete the association files created when a venv was associated with a script."""
        if not script_dir or not os.path.exists(script_dir):
            if show_results:
                return {"deleted": [], "missing": [], "failed": ["Directory not found"]}
            return
            
        files_to_delete = [
            ".venv-association",
            "run_this.vbs",
            "run_this.bat",
            "launcher.py",
            "launcher_icon.ico"
        ]
        
        result = {"deleted": [], "missing": [], "failed": []}
        any_files_found = False
        
        for file_name in files_to_delete:
            file_path = os.path.join(script_dir, file_name)
            if os.path.exists(file_path):
                any_files_found = True
                try:
                    os.remove(file_path)
                    print(f"Deleted association file: {file_path}")
                    if show_results:
                        result["deleted"].append(file_name)
                except Exception as e:
                    print(f"Error deleting {file_path}: {e}")
                    if show_results:
                        result["failed"].append(f"{file_name} ({str(e)})")
            elif show_results:
                result["missing"].append(file_name)
        
        # Special note if none of the expected files were found
        if show_results and not any_files_found:
            result["missing"] = ["No association files were found in the selected directory"]
            
        return result if show_results else None

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

        # Check if we're updating an existing venv or creating a new one
        if hasattr(self, 'selected_venv_path') and self.selected_venv_path:
            self.update_existing_venv(self.selected_venv_path)
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
                
                # Always add PySide6 and customtkinter to requirements
                if "PySide6" not in requirements:
                    requirements.append("PySide6")
                if "customtkinter" not in requirements:
                    requirements.append("customtkinter")
                
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
            except Exception as e:
                self.hide_spinner()
                self.show_error(f"Error creating virtual environment: {str(e)}")
                
        threading.Thread(target=venv_task, daemon=True).start()

    def show_error(self, message):
        """Show an error message with safe window handling"""
        dialog = ctk.CTkToplevel(self)
        dialog.title("Error")
        dialog.geometry("400x150")
        dialog.transient(self)
        dialog.grab_set()
        
        # Safe focus approach
        if self.winfo_exists():
            dialog.lift()
            if dialog.winfo_exists():
                dialog.focus_set()
                
        # Center the dialog only if it still exists
        if dialog.winfo_exists():
            self.center_dialog(dialog)
            
        ctk.CTkLabel(dialog, text=message, wraplength=350).pack(pady=20, padx=20)
        ctk.CTkButton(dialog, text="OK", command=dialog.destroy).pack(pady=10)

    def show_success(self, message):
        """Show a success message with safe window handling"""
        dialog = ctk.CTkToplevel(self)
        dialog.title("Success")
        dialog.geometry("400x200")  # Increased height to accommodate potential button
        dialog.transient(self)
        dialog.grab_set()
        
        # Safe focus approach
        if self.winfo_exists():
            dialog.lift()
            if dialog.winfo_exists():
                dialog.focus_set()
                
        # Center the dialog only if it still exists
        if dialog.winfo_exists():
            self.center_dialog(dialog)
        
        # Check if this is a message about deleting the app's venv
        is_app_venv_message = "You have deleted the virtual environment used by this application" in message
        
        ctk.CTkLabel(dialog, text=message, wraplength=350).pack(pady=20, padx=20)
        
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(pady=10)
        
        if is_app_venv_message:
            # Add a button to reinstall packages
            ctk.CTkButton(btn_frame, text="Reinstall Packages", fg_color="#1A73E8", hover_color="#1557B0", 
                       command=lambda: [dialog.destroy(), self.reinstall_app_packages()]).pack(side="left", padx=10)
            ctk.CTkButton(btn_frame, text="OK", command=dialog.destroy).pack(side="right", padx=10)
        else:
            ctk.CTkButton(btn_frame, text="OK", command=dialog.destroy).pack(padx=10)

    def reinstall_app_packages(self):
        """Reinstall packages required by this application"""
        req_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "requirements.txt")
        if not os.path.exists(req_file):
            self.show_error("requirements.txt not found. Cannot reinstall packages.")
            return
            
        def install_task():
            self.show_spinner("Reinstalling required packages...")
            try:
                # Use pip to reinstall the packages
                cmd = [sys.executable, "-m", "pip", "install", "-r", req_file]
                result = subprocess.run(cmd, capture_output=True, text=True)
                
                if result.returncode == 0:
                    self.hide_spinner()
                    self.show_success("Packages successfully reinstalled. You may need to restart the application.")
                else:
                    self.hide_spinner()
                    error_msg = f"Error reinstalling packages:\n{result.stderr}"
                    self.show_error(error_msg)
            except Exception as e:
                self.hide_spinner()
                self.show_error(f"Error reinstalling packages: {str(e)}")
                
        threading.Thread(target=install_task, daemon=True).start()

    def show_install_summary(self, success, failed, skipped):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Venv Creation Summary")
        
        # Configure but don't show yet
        dialog.geometry("450x350")
        dialog.transient(self)
        dialog.withdraw()  # Hide initially
        
        # Set icon if available
        if self.icon_path and os.path.exists(self.icon_path):
            try:
                dialog.iconbitmap(self.icon_path)
            except Exception:
                pass
        
        # Create content
        msg = ""
        
        # First, identify if PySide6 or customtkinter were installed and remove them from the regular success list
        auto_installed = []
        for pkg in ["PySide6", "customtkinter"]:
            if pkg in success:
                success.remove(pkg)
                auto_installed.append(pkg)
        
        # Show auto-installed packages first
        if auto_installed:
            msg += "⚙️ Automatically installed packages:\n" + ", ".join(auto_installed) + " (required for UI and system tray support)\n\n"
        
        # Then show regular packages
        if success:
            msg += "✔️ Packages successfully installed:\n" + "\n".join(success) + "\n\n"
        if failed:
            msg += "⛔ Packages failed to install:\n" + "\n".join(failed) + "\n\n"
        if skipped:
            msg += "ℹ️ Skipped standard library modules:\n" + "\n".join(skipped) + "\n\n"
        if not (success or auto_installed or failed or skipped):
            msg = "No additional packages were specified."
            
        ctk.CTkLabel(dialog, text=msg, wraplength=400, justify="left").pack(pady=20, padx=20)
        ctk.CTkButton(dialog, text="OK", command=dialog.destroy).pack(pady=10)
        
        # Center and show the dialog
        self.center_dialog(dialog)
        dialog.grab_set()
        dialog.deiconify()  # Now show the dialog

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
        # Get the target directory from the first file path
        target_dir = os.path.dirname(file_paths[0])
        def scrape_task():
            self.show_spinner("Scraping requirements from selected files...")
            try:
                # Pass the file paths to requirements_generator.py (updated to match new arg format)
                cmd = [sys.executable, script_path] + list(file_paths)
                result = subprocess.run(cmd, capture_output=True, text=True)
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
            "type": "venv",
            "date": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "parent": os.path.basename(parent_dir),
            "venv_path": venv_path
        }
        try:
            if os.path.exists(ENV_HISTORY_FILE):
                with open(ENV_HISTORY_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
            else:
                data = []
            data.append(entry)
            with open(ENV_HISTORY_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            print(f"Error logging venv creation: {e}")

    def log_venv_deletion(self, venv_path):
        """Mark a venv as deleted in the history file."""
        try:
            if os.path.exists(ENV_HISTORY_FILE):
                with open(ENV_HISTORY_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                
                # Find the corresponding entry and mark it as deleted
                for entry in data:
                    if entry.get("type") == "venv" and entry.get("venv_path") == venv_path:
                        entry["date_deleted"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                # Save the updated data
                with open(ENV_HISTORY_FILE, "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=2)
        except Exception as e:
            print(f"Error logging venv deletion: {e}")

    def open_venv_history(self):
        # Dialog for environment history
        dialog = ctk.CTkToplevel(self)
        dialog.title("Environment History")
        dialog.geometry("1300x500")  # Increased width from 1080 to 1300 (about 20% wider)
        dialog.transient(self)
        dialog.grab_set()
        dialog.lift()
        dialog.focus_set()
        self.center_dialog(dialog)
        
        # Top controls frame
        controls_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        controls_frame.pack(fill="x", padx=20, pady=(10, 0))
        
        # Hide deleted items checkbox
        hide_deleted_var = tk.BooleanVar(value=False)
        hide_cb = ctk.CTkCheckBox(controls_frame, text="Hide deleted items", variable=hide_deleted_var)
        hide_cb.pack(side="left", pady=(5, 5))
        
        # Type filter
        type_frame = ctk.CTkFrame(controls_frame, fg_color="transparent")
        type_frame.pack(side="right")
        
        ctk.CTkLabel(type_frame, text="Show:").pack(side="left", padx=(0, 5))
        
        show_type_var = tk.StringVar(value="all")

        # Table frame
        table_frame = ctk.CTkFrame(dialog, fg_color="#23272E")
        table_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Sorting state
        sort_state = {"column": None, "reverse": False}

        # Header row with clickable headers for sorting
        headers = [
            ("Type", 100),            # Increased from 80
            ("Date Added", 130),      # Increased from 120
            ("Last Used", 220),         # Changed from "Details"
            ("Date Removed", 130),    # Increased from 120
            ("Path", 450),            # Increased from 350
            ("Actions", 180)          # Increased from 120 for button area
        ]
        
        def sort_by_column(col_idx):
            nonlocal sort_state
            
            # If clicking the same column, toggle sort direction
            if sort_state["column"] == col_idx:
                sort_state["reverse"] = not sort_state["reverse"]
            else:
                sort_state["column"] = col_idx
                sort_state["reverse"] = False
                
            refresh_history()
        
        for col, (header, width) in enumerate(headers):
            if col < len(headers) - 1:  # Don't make the last column (buttons) sortable
                header_label = ctk.CTkButton(
                    table_frame, 
                    text=header, 
                    font=("Arial", 12, "bold"),
                    width=width,
                    anchor="w",
                    fg_color="transparent",
                    hover_color="#333740",
                    command=lambda c=col: sort_by_column(c)
                )
                header_label.grid(row=0, column=col, padx=5, pady=(0, 4), sticky="w")
            else:
                ctk.CTkLabel(table_frame, text=header, font=("Arial", 12, "bold"), width=width, anchor="w").grid(row=0, column=col, padx=5, pady=(0, 4), sticky="w")

        def get_venv_last_used_date(venv_path):
            """Get the last modified time of the activate script in the venv"""
            if not venv_path or not os.path.exists(venv_path):
                return ""
                
            # Path to activate script
            activate_path = os.path.join(venv_path, "Scripts" if sys.platform == "win32" else "bin", "activate")
            
            # Check if activate script exists
            if os.path.exists(activate_path):
                try:
                    # Get the modification time
                    mod_time = os.path.getmtime(activate_path)
                    # Convert to datetime and format
                    return datetime.datetime.fromtimestamp(mod_time).strftime("%Y-%m-%d %H:%M:%S")
                except Exception:
                    pass
                    
            return ""
            
        def path_exists(path):
            """Check if a path exists"""
            return path and os.path.exists(path)

        def refresh_history():
            # Remove all rows except header
            for widget in table_frame.winfo_children():
                info = widget.grid_info()
                if info and 'row' in info and info['row'] != 0:
                    widget.destroy()
            
            # Load data
            try:
                if os.path.exists(ENV_HISTORY_FILE):
                    with open(ENV_HISTORY_FILE, "r", encoding="utf-8") as f:
                        data = json.load(f)
                else:
                    data = []
            except Exception as e:
                print(f"Error loading environment history: {e}")
                data = []
            
            # Apply filters
            filtered_data = []
            for entry in data:
                # Check for deleted items filter
                if hide_deleted_var.get() and entry.get("date_deleted"):
                    continue
                
                # Check for type filter
                entry_type = entry.get("type", "venv")  # Default to "venv" for backward compatibility
                if show_type_var.get() != "all" and entry_type != show_type_var.get():
                    continue
                
                filtered_data.append(entry)
                
            # Add dynamic data for sorting and display
            for entry in filtered_data:
                entry_type = entry.get("type", "venv")
                
                # Different behavior based on entry type
                if entry_type == "venv":
                    # Get last used date from file modification time
                    entry["_details"] = get_venv_last_used_date(entry.get("venv_path", ""))
                    # Check if venv directory actually exists
                    entry["_path_exists"] = path_exists(entry.get("venv_path", ""))
                    # Set display path
                    entry["_display_path"] = entry.get("venv_path", "")
                    
                elif entry_type == "association":
                    # For associations, show script name and last modified
                    script_name = entry.get("script_name", "")
                    script_mod = entry.get("script_last_modified", "")
                    entry["_details"] = f"Script: {script_name}\nLast Modified: {script_mod}"
                    
                    # Check if any association files exist in the script directory
                    script_dir = entry.get("script_dir", "")
                    association_files_exist = False
                    
                    if os.path.exists(script_dir):
                        association_files = [
                            ".venv-association",
                            "run_this.vbs",
                            "run_this.bat",
                            "launcher.py",
                            "launcher_icon.ico"
                        ]
                        
                        for file_name in association_files:
                            if os.path.exists(os.path.join(script_dir, file_name)):
                                association_files_exist = True
                                break
                    
                    # Update path exists flag based on actual file check
                    entry["_path_exists"] = association_files_exist
                    
                    # Set display path to show both script and venv paths
                    script_path = entry.get("script_path", "")
                    venv_path = entry.get("venv_path", "")
                    entry["_display_path"] = f"Script: {script_path}\nVenv: {venv_path}"
                
            # Apply sorting if a column is selected
            if sort_state["column"] is not None:
                col_idx = sort_state["column"]
                reverse = sort_state["reverse"]
                
                # Map column index to the appropriate sorting key
                sort_keys = {
                    0: lambda e: e.get("type", ""),  # Type
                    1: lambda e: e.get("date", ""),  # Date Added
                    2: lambda e: e.get("_details", ""),  # Last Used
                    3: lambda e: e.get("date_deleted", ""),  # Date Removed
                    4: lambda e: e.get("_display_path", "").lower(),  # Path
                }
                
                if col_idx in sort_keys:
                    filtered_data.sort(key=sort_keys[col_idx], reverse=reverse)
            
            for i, entry in enumerate(filtered_data):
                row = i + 1  # header is row 0
                entry_type = entry.get("type", "venv")
                date_added = entry.get("date", "")
                details = entry.get("_details", "")
                date_removed = entry.get("date_deleted", "")
                display_path = entry.get("_display_path", "")
                path_exists_flag = entry.get("_path_exists", False)
                
                # Type column (with icon)
                type_text = "Virtual Env" if entry_type == "venv" else "Script Association"
                type_label = ctk.CTkLabel(table_frame, text=type_text, width=80, anchor="w")
                type_label.grid(row=row, column=0, padx=5, sticky="w")
                
                # Date added
                ctk.CTkLabel(table_frame, text=date_added, width=120, anchor="w").grid(row=row, column=1, padx=5, sticky="w")
                
                # Last Used column (previously Details)
                details_label = ctk.CTkLabel(table_frame, text=details, width=200, anchor="w")
                details_label.grid(row=row, column=2, padx=5, sticky="w")
                
                # Date removed
                date_text = date_removed
                date_color = "white"
                
                # Add manual deletion indicator if resources don't exist but aren't marked as deleted
                if not path_exists_flag and not date_removed:
                    date_text = "Manual Delete Detected"
                    date_color = "#FF6B6B"  # Bright red for warning
                
                date_label = ctk.CTkLabel(table_frame, text=date_text, width=120, anchor="w", text_color=date_color)
                date_label.grid(row=row, column=3, padx=5, sticky="w")
                
                # Path column
                path_color = "#1A73E8" if path_exists_flag else "#e74c3c"
                
                def open_containing_folder(path):
                    if entry_type == "association":
                        # For associations, use the script_dir
                        folder = entry.get("script_dir", "")
                    elif os.path.isfile(path):
                        folder = os.path.dirname(path)
                    else:
                        folder = path
                    
                    if os.path.exists(folder):
                        if platform.system() == "Windows":
                            os.startfile(folder)
                        elif platform.system() == "Darwin":
                            subprocess.Popen(["open", folder])
                        else:
                            subprocess.Popen(["xdg-open", folder])
                    else:
                        self.show_error(f"Folder does not exist: {folder}")
                
                def make_link_callback(entry=entry):
                    return lambda e: open_containing_folder(entry.get("_display_path", ""))
                    
                path_label = ctk.CTkLabel(table_frame, text=display_path, width=450, anchor="w", 
                                       text_color=path_color, cursor="hand2")
                path_label.grid(row=row, column=4, padx=5, sticky="w")
                path_label.bind("<Button-1>", make_link_callback())
                
                # Create action buttons frame
                btn_frame = ctk.CTkFrame(table_frame, fg_color="transparent")
                btn_frame.grid(row=row, column=5, padx=5, sticky="w")
                
                # Resource exists check
                resource_exists = False
                
                if entry_type == "venv":
                    # Check if the venv directory actually exists
                    venv_path = entry.get("venv_path", "")
                    resource_exists = os.path.exists(venv_path)
                else:  # association
                    # Check if any association files exist in the script directory
                    script_dir = entry.get("script_dir", "")
                    if os.path.exists(script_dir):
                        association_files = [
                            ".venv-association",
                            "run_this.vbs",
                            "run_this.bat",
                            "launcher.py",
                            "launcher_icon.ico"
                        ]
                        
                        for file_name in association_files:
                            file_path = os.path.join(script_dir, file_name)
                            if os.path.exists(file_path):
                                resource_exists = True
                                break
                
                # Always allow delete entry if the item has already been marked as deleted
                if date_removed:
                    resource_exists = False
                
                # Delete resource buttons
                if resource_exists and not date_removed:
                    if entry_type == "venv":
                        # Venv delete button
                        venv_btn = ctk.CTkButton(
                            btn_frame, 
                            text="Delete Venv", 
                            width=70,  # Wider button
                            height=24,
                            fg_color="#e74c3c", 
                            hover_color="#b93a2b", 
                            command=lambda idx=i, path=entry.get("venv_path", ""): 
                                self.delete_venv_from_history(idx, path, dialog, refresh_history)
                        )
                        venv_btn.pack(side="left", padx=(0, 4))
                    else:
                        # Script association delete button
                        files_btn = ctk.CTkButton(
                            btn_frame, 
                            text="Delete Files", 
                            width=70,  # Wider button
                            height=24,
                            fg_color="#e74c3c", 
                            hover_color="#b93a2b", 
                            command=lambda idx=i, dir=entry.get("script_dir", ""): 
                                self.delete_association_from_history(dir, dialog, refresh_history, idx)
                        )
                        files_btn.pack(side="left", padx=(0, 4))
                
                # Delete Entry button (disabled if resources exist)
                entry_btn = ctk.CTkButton(
                    btn_frame, 
                    text="Delete Entry", 
                    width=70,  # Wider button
                    height=24,
                    fg_color="#FF8C00" if not resource_exists else "#A9A9A9", 
                    hover_color="#E57C00" if not resource_exists else "#A9A9A9", 
                    command=lambda idx=i, e=entry: self.remove_history_entry(idx, dialog, refresh_history, e),
                    state="normal" if not resource_exists else "disabled"
                )
                entry_btn.pack(side="left")
                
                # Add right-click context menu
                def show_context_menu(event, idx=i):
                    # Create a context menu
                    context_menu = tk.Menu(dialog, tearoff=0)
                    
                    # Copy path options
                    if entry_type == "venv":
                        venv_path = entry.get("venv_path", "")
                        context_menu.add_command(
                            label="Copy Venv Path", 
                            command=lambda: self.copy_to_clipboard(venv_path)
                        )
                    else:  # association
                        script_path = entry.get("script_path", "")
                        script_dir = entry.get("script_dir", "")
                        venv_path = entry.get("venv_path", "")
                        
                        context_menu.add_command(
                            label="Copy Script Path", 
                            command=lambda: self.copy_to_clipboard(script_path)
                        )
                        context_menu.add_command(
                            label="Copy Script Directory", 
                            command=lambda: self.copy_to_clipboard(script_dir)
                        )
                        context_menu.add_command(
                            label="Copy Venv Path", 
                            command=lambda: self.copy_to_clipboard(venv_path)
                        )
                    
                    context_menu.add_separator()
                    
                    # Force delete option
                    context_menu.add_command(
                        label="Force Delete Entry", 
                        command=lambda: self.force_delete_entry(idx, dialog, refresh_history)
                    )
                    
                    # Display the menu
                    try:
                        context_menu.tk_popup(event.x_root, event.y_root)
                    finally:
                        context_menu.grab_release()
                
                # Bind the context menu to the whole row for right-click
                for col in range(6):  # All columns
                    widget = table_frame.grid_slaves(row=row, column=col)[0]
                    widget.bind("<Button-3>", show_context_menu)
                
                # Add tooltip for "Force Delete"
                try:
                    from CTkToolTip import CTkToolTip
                    if resource_exists:
                        CTkToolTip(entry_btn, message="Delete the resources first before removing from history, or right-click and select 'Force Delete Entry'")
                    else:
                        CTkToolTip(entry_btn, message="Click to remove this entry from history")
                except Exception:
                    pass
        
        # Create radio buttons now that refresh_history is defined
        def create_radio_button(value, text):
            return ctk.CTkRadioButton(
                type_frame, 
                text=text, 
                variable=show_type_var, 
                value=value, 
                command=refresh_history
            )
        
        all_radio = create_radio_button("all", "All")
        venv_radio = create_radio_button("venv", "Virtual Envs")
        assoc_radio = create_radio_button("association", "Script Associations")
        
        all_radio.pack(side="left", padx=5)
        venv_radio.pack(side="left", padx=5)
        assoc_radio.pack(side="left", padx=5)
                
        # Button frame for bottom actions
        bottom_btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        bottom_btn_frame.pack(fill="x", padx=10, pady=(5, 10))
        
        # Export button (bottom left)
        export_btn = ctk.CTkButton(
            bottom_btn_frame, 
            text="Export History", 
            width=120, 
            height=32,
            fg_color="#1565c0", 
            hover_color="#0d47a1", 
            command=lambda: self.export_venv_history(dialog)
        )
        export_btn.pack(side="left", padx=5, pady=5)
        
        # View History button (bottom left, next to export)
        view_btn = ctk.CTkButton(
            bottom_btn_frame, 
            text="View History", 
            width=120, 
            height=32,
            fg_color="#8E44AD",  # Purple
            hover_color="#6C3483", 
            command=lambda: self.view_history_file(dialog)
        )
        view_btn.pack(side="left", padx=5, pady=5)
        
        # Button dimensions for right-side buttons
        btn_width = 180
        btn_height = 32
        btn_spacing = 10
        
        # Browse To Remove association button (purple)
        browse_remove_assoc_btn = ctk.CTkButton(
            dialog, 
            text="Delete Script Association", 
            width=btn_width, 
            height=btn_height, 
            fg_color="#8E44AD",  # Purple 
            hover_color="#6C3483", 
            command=lambda: self.browse_remove_association(dialog, refresh_history)
        )
        browse_remove_assoc_btn.place(relx=1.0, rely=1.0, x=-20, y=-20-btn_height-btn_spacing-btn_height-btn_spacing, anchor="se")
        
        # Browse To Remove venv button (blue)
        browse_remove_btn = ctk.CTkButton(
            dialog, 
            text="Delete Virtual Environment", 
            width=btn_width, 
            height=btn_height, 
            fg_color="#1565c0",  # Blue
            hover_color="#0d47a1", 
            command=lambda: self.browse_remove_venv(dialog, refresh_history)
        )
        browse_remove_btn.place(relx=1.0, rely=1.0, x=-20, y=-20-btn_height-btn_spacing, anchor="se")
        
        # Close button (red, at bottom)
        close_btn = ctk.CTkButton(
            dialog, 
            text="Close", 
            width=btn_width, 
            height=btn_height, 
            fg_color="#e74c3c",  # Red
            hover_color="#b93a2b", 
            command=dialog.destroy
        )
        close_btn.place(relx=1.0, rely=1.0, x=-20, y=-20, anchor="se")
        
        hide_cb.configure(command=refresh_history)
        refresh_history()

    def remove_history_entry(self, idx, dialog, refresh_history, entry):
        """Remove an entry from the history."""
        if not os.path.exists(ENV_HISTORY_FILE):
            self.show_error("History file not found.")
            return
        
        try:
            # Load the current data
            with open(ENV_HISTORY_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                
            if 0 <= idx < len(data):
                entry_type = entry.get("type", "venv")
                date_deleted = entry.get("date_deleted", "")
                
                # Skip resource existence check if already marked as deleted
                if not date_deleted:
                    # Direct physical check for resource existence
                    resource_exists = False
                    
                    if entry_type == "venv":
                        venv_path = entry.get("venv_path", "")
                        # Check if the venv directory still exists
                        resource_exists = os.path.exists(venv_path)
                    else:  # association
                        script_dir = entry.get("script_dir", "")
                        # Check if any association files exist
                        if os.path.exists(script_dir):
                            association_files = [
                                ".venv-association",
                                "run_this.vbs",
                                "run_this.bat",
                                "launcher.py",
                                "launcher_icon.ico"
                            ]
                            
                            for file_name in association_files:
                                if os.path.exists(os.path.join(script_dir, file_name)):
                                    resource_exists = True
                                    break
                    
                    # If resources still exist, show error message
                    if resource_exists:
                        self.show_error("Cannot remove history entry while physical resources still exist.\n\nDelete the resources first with the Delete button.")
                        return
                
                # Simplified confirmation dialog - just ask about removing from history
                if entry_type == "venv":
                    message_main = "Remove this virtual environment from the Environment History list?"
                else:
                    message_main = "Remove this script association from the Environment History list?"
                
                message_detail = "This will permanently remove the entry from history."
                
                confirm = ConfirmationDialog(
                    self, "Confirm Removal", message_main, message_detail,
                    "Remove Entry", "Cancel", icon_path=self.icon_path
                )
                
                # Store reference to dialog
                confirm_ref = {"dialog": confirm}
                
                # Set action handlers
                confirm.ok_action = lambda: self.delete_entry_from_json(idx, dialog, refresh_history, confirm_ref["dialog"])
                confirm.cancel_action = lambda: confirm.destroy()
                
                if self.winfo_exists() and confirm.winfo_exists():
                    self.center_dialog(confirm, parent=dialog)
            else:
                self.show_error(f"Invalid index for history entry: {idx}")
        except Exception as e:
            message_title = "Error"
            message_detail = str(e)
            message_main = "Error removing history entry."
            confirm = ConfirmationDialog(
                self, message_title, message_main, message_detail,
                "OK", "Cancel", icon_path=self.icon_path
            )
            
            # Set action handlers to ensure the dialog is destroyed
            confirm.ok_action = lambda: confirm.destroy()
            confirm.cancel_action = lambda: confirm.destroy()
            
            if self.winfo_exists() and confirm.winfo_exists():
                self.center_dialog(confirm, parent=dialog)

    def delete_entry_from_json(self, idx, dialog, refresh_history, confirm_dialog=None):
        """Remove an entry completely from the JSON history file."""
        try:
            if os.path.exists(ENV_HISTORY_FILE):
                with open(ENV_HISTORY_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                if 0 <= idx < len(data):
                    # Get the entry for success message
                    entry = data[idx]
                    entry_type = entry.get("type", "venv")
                    
                    # Remove the entry
                    del data[idx]
                    
                    # Save the modified data
                    with open(ENV_HISTORY_FILE, 'w', encoding='utf-8') as f:
                        json.dump(data, f, indent=2)
                    
                    # First close the confirmation dialog if it exists
                    if confirm_dialog and confirm_dialog.winfo_exists():
                        confirm_dialog.destroy()
                        
                    # Show success message
                    if entry_type == "venv":
                        self.show_success(f"Successfully removed virtual environment entry from the Environment History list.")
                    else:
                        self.show_success(f"Successfully removed script association entry from the Environment History list.")
                    
                    # Refresh the display
                    refresh_history()
                else:
                    # Close confirmation dialog first
                    if confirm_dialog and confirm_dialog.winfo_exists():
                        confirm_dialog.destroy()
                    self.show_error(f"Invalid index for history entry: {idx}")
            else:
                # Close confirmation dialog first
                if confirm_dialog and confirm_dialog.winfo_exists():
                    confirm_dialog.destroy()
                self.show_error("History file not found.")
        except Exception as e:
            # Close confirmation dialog first
            if confirm_dialog and confirm_dialog.winfo_exists():
                confirm_dialog.destroy()
            self.show_error(f"Error updating history: {str(e)}")
            if refresh_history:
                refresh_history()  # Refresh anyway to maintain UI consistency
                
    def delete_association_from_history(self, script_dir, dialog, refresh_history, idx=None):
        """Delete association files and mark the entry as deleted in history."""
        try:
            # Delete the association files
            result = self.delete_association_files(script_dir, show_results=True)
            
            # If an index is provided, always mark the entry as deleted in the history
            if idx is not None:
                try:
                    if os.path.exists(ENV_HISTORY_FILE):
                        with open(ENV_HISTORY_FILE, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                        
                        if 0 <= idx < len(data):
                            # Mark as deleted
                            data[idx]["date_deleted"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            
                            # Save the modified data
                            with open(ENV_HISTORY_FILE, 'w', encoding='utf-8') as f:
                                json.dump(data, f, indent=2)
                except Exception as e:
                    print(f"Error marking association as deleted: {e}")
                    error_dialog = ConfirmationDialog(
                        self, "Error", "Could not mark association as deleted in history",
                        str(e), "OK", "Cancel", icon_path=self.icon_path
                    )
                    error_dialog.ok_action = lambda: error_dialog.destroy()
                    error_dialog.cancel_action = lambda: error_dialog.destroy()
                    
                    if self.winfo_exists() and error_dialog.winfo_exists():
                        self.center_dialog(error_dialog, parent=dialog)
            
            # Show success message
            if result["deleted"]:
                message = "Successfully deleted the following files:\n• " + "\n• ".join(result["deleted"])
                if result["missing"]:
                    message += "\n\nThe following files were not found:\n• " + "\n• ".join(result["missing"])
                if result["failed"]:
                    message += "\n\nFailed to delete the following files:\n• " + "\n• ".join(result["failed"])
                self.show_success(message)
            else:
                if result["missing"]:
                    self.show_error(f"No association files were found in:\n{script_dir}")
                elif result["failed"]:
                    self.show_error(f"Failed to delete association files:\n• " + "\n• ".join(result["failed"]))
            
            # Always refresh the history display
            refresh_history()
            
        except Exception as e:
            error_dialog = ConfirmationDialog(
                self, "Error", "Could not delete association files",
                str(e), "OK", "Cancel", icon_path=self.icon_path
            )
            error_dialog.ok_action = lambda: (error_dialog.destroy(), refresh_history())
            error_dialog.cancel_action = lambda: error_dialog.destroy()
            
            if self.winfo_exists() and error_dialog.winfo_exists():
                self.center_dialog(error_dialog, parent=dialog)

    def browse_remove_association(self, dialog, refresh_history):
        """Browse for a directory containing association files to remove."""
        # Open a directory selection dialog
        script_dir = filedialog.askdirectory(title="Select Directory with Script Association Files")
        if not script_dir:
            return
        
        # Check if any association files exist
        association_file = os.path.join(script_dir, ".venv-association")
        any_files_exist = os.path.exists(association_file)
        
        if not any_files_exist:
            # Double-check for other files
            for file_name in ["run_this.vbs", "run_this.bat", "launcher.py", "launcher_icon.ico"]:
                if os.path.exists(os.path.join(script_dir, file_name)):
                    any_files_exist = True
                    break
        
        if not any_files_exist:
            self.show_error(f"No association files were found in:\n{script_dir}")
            return
        
        # Ask for confirmation
        message_main = "Are you sure you want to delete the association files?"
        message_detail = f"This will delete any association files found in:\n{script_dir}\n\nThis action cannot be undone."
        
        confirm = ConfirmationDialog(
            self, "Confirm Deletion", message_main, message_detail,
            "Delete Files", "Cancel", icon_path=self.icon_path
        )
        
        def on_ok():
            if confirm.winfo_exists():
                confirm.destroy()
            self.delete_association_from_history(script_dir, dialog, refresh_history)
        
        def on_cancel():
            if confirm.winfo_exists():
                confirm.destroy()
        
        confirm.ok_action = on_ok
        confirm.cancel_action = on_cancel
        
        if self.winfo_exists() and confirm.winfo_exists():
            self.center_dialog(confirm, parent=dialog)

    def export_venv_history(self, parent_dialog):
        """Export environment history to a text file"""
        file_path = filedialog.asksaveasfilename(
            parent=parent_dialog,
            title="Export Environment History", 
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
            
        try:
            if os.path.exists(ENV_HISTORY_FILE):
                with open(ENV_HISTORY_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
            else:
                data = []
                
            with open(file_path, "w", encoding="utf-8") as f:
                f.write("Environment History Log\n")
                f.write("============================\n\n")
                
                for i, entry in enumerate(data):
                    entry_type = entry.get("type", "venv")
                    date_added = entry.get("date", "N/A")
                    date_removed = entry.get("date_deleted", "")
                    
                    f.write(f"Entry #{i+1} ({entry_type}):\n")
                    f.write(f"  Date Added: {date_added}\n")
                    f.write(f"  Date Removed: {date_removed if date_removed else 'N/A'}\n")
                    
                    if entry_type == "venv":
                        script_root = entry.get("parent", "")
                        venv_path = entry.get("venv_path", "")
                        last_used = get_venv_last_used_date(venv_path)
                        exists = "Yes" if os.path.exists(venv_path) else "No"
                        
                        f.write(f"  Last Used: {last_used if last_used else 'Never/Unknown'}\n")
                        f.write(f"  Script Root: {script_root}\n")
                        f.write(f"  Venv Path: {venv_path}\n")
                        f.write(f"  Still Exists: {exists}\n")
                    
                    elif entry_type == "association":
                        script_path = entry.get("script_path", "")
                        script_name = entry.get("script_name", "")
                        script_dir = entry.get("script_dir", "")
                        venv_path = entry.get("venv_path", "")
                        script_mod = entry.get("script_last_modified", "")
                        systray = "Yes" if entry.get("with_systray", False) else "No"
                        startup = "Yes" if entry.get("with_startup", False) else "No"
                        exists = "Yes" if os.path.exists(script_path) else "No"
                        
                        f.write(f"  Script Name: {script_name}\n")
                        f.write(f"  Script Path: {script_path}\n")
                        f.write(f"  Script Directory: {script_dir}\n")
                        f.write(f"  Venv Path: {venv_path}\n")
                        f.write(f"  Last Modified: {script_mod}\n")
                        f.write(f"  With System Tray: {systray}\n")
                        f.write(f"  With Startup: {startup}\n")
                        f.write(f"  Script Still Exists: {exists}\n")
                    
                    f.write("\n")
                    
            self.show_success(f"Exported environment history to {file_path}")
            
        except Exception as e:
            self.show_error(f"Error exporting history: {str(e)}")
            
    def view_history_file(self, parent_dialog):
        """Open a dialog to view the history file contents in a human-readable format"""
        try:
            # Check if history file exists
            if not os.path.exists(ENV_HISTORY_FILE):
                self.show_error("History file does not exist yet.")
                return
                
            # Create a new dialog to display the history content
            viewer = ctk.CTkToplevel(self)
            viewer.title("Environment History Details")
            viewer.geometry("800x600")
            viewer.transient(parent_dialog)
            viewer.grab_set()
            viewer.lift()
            
            # Center the dialog
            self.center_dialog(viewer, parent=parent_dialog)
            
            # Create a frame for the text widget
            frame = ctk.CTkFrame(viewer)
            frame.pack(fill="both", expand=True, padx=10, pady=10)
            
            # Create a text widget with scrollbar
            text_widget = tk.Text(
                frame, 
                wrap="word", 
                bg="#232830", 
                fg="white",
                font=("Consolas", 11),
                padx=10,
                pady=10
            )
            
            # Add scrollbars
            y_scrollbar = tk.Scrollbar(frame, orient="vertical", command=text_widget.yview)
            x_scrollbar = tk.Scrollbar(frame, orient="horizontal", command=text_widget.xview)
            text_widget.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
            
            # Pack scrollbars and text widget
            y_scrollbar.pack(side="right", fill="y")
            x_scrollbar.pack(side="bottom", fill="x")
            text_widget.pack(side="left", fill="both", expand=True)
            
            # Load and format the history file contents
            with open(ENV_HISTORY_FILE, "r", encoding="utf-8") as f:
                try:
                    data = json.load(f)
                    
                    # Format the data in a human-readable way
                    text_widget.tag_configure("title", font=("Arial", 12, "bold"), foreground="#64B5F6")
                    text_widget.tag_configure("heading", font=("Arial", 11, "bold"), foreground="#81C784")
                    text_widget.tag_configure("subheading", font=("Arial", 10, "bold"), foreground="#FFB74D")
                    text_widget.tag_configure("normal", font=("Consolas", 10), foreground="white")
                    text_widget.tag_configure("deleted", font=("Consolas", 10), foreground="#E57373")
                    text_widget.tag_configure("path", font=("Consolas", 10), foreground="#AED581")
                    text_widget.tag_configure("timestamp", font=("Consolas", 10), foreground="#90CAF9")
                    text_widget.tag_configure("separator", font=("Arial", 10), foreground="#555555")
                    
                    # Add title
                    text_widget.insert("end", "Environment History\n\n", "title")
                    
                    # Count entries by type
                    venv_count = len([e for e in data if e.get("type", "venv") == "venv"])
                    assoc_count = len([e for e in data if e.get("type", "") == "association"])
                    deleted_count = len([e for e in data if e.get("date_deleted", "")])
                    
                    # Add summary
                    text_widget.insert("end", f"Total Entries: {len(data)}\n", "heading")
                    text_widget.insert("end", f"Virtual Environments: {venv_count}\n", "normal")
                    text_widget.insert("end", f"Script Associations: {assoc_count}\n", "normal")
                    text_widget.insert("end", f"Deleted Items: {deleted_count}\n\n", "normal")
                    
                    # Add separator
                    text_widget.insert("end", "=" * 80 + "\n\n", "separator")
                    
                    # Add each entry
                    for i, entry in enumerate(data):
                        entry_type = entry.get("type", "venv")
                        date_added = entry.get("date", "Unknown")
                        date_deleted = entry.get("date_deleted", "")
                        
                        # Entry header
                        header = f"Entry #{i+1}: "
                        if entry_type == "venv":
                            header += "Virtual Environment"
                        else:
                            header += "Script Association"
                            
                        if date_deleted:
                            header += " [DELETED]"
                            
                        text_widget.insert("end", header + "\n", "heading")
                        
                        # Common fields
                        text_widget.insert("end", "Date Created: ", "subheading")
                        text_widget.insert("end", f"{date_added}\n", "timestamp")
                        
                        if date_deleted:
                            text_widget.insert("end", "Date Deleted: ", "subheading")
                            text_widget.insert("end", f"{date_deleted}\n", "deleted")
                        
                        # Type-specific fields
                        if entry_type == "venv":
                            venv_path = entry.get("venv_path", "")
                            parent_dir = entry.get("parent", "")
                            
                            text_widget.insert("end", "Parent Directory: ", "subheading")
                            text_widget.insert("end", f"{parent_dir}\n", "normal")
                            
                            text_widget.insert("end", "Virtual Environment Path: ", "subheading")
                            text_widget.insert("end", f"{venv_path}\n", "path")
                            
                            # Check if it exists
                            exists = os.path.exists(venv_path)
                            text_widget.insert("end", "Still Exists: ", "subheading")
                            if exists:
                                text_widget.insert("end", "Yes\n", "normal")
                            else:
                                if date_deleted:
                                    text_widget.insert("end", "No (Deleted through app)\n", "deleted")
                                else:
                                    text_widget.insert("end", "No (Manual Delete Detected Outside of App)\n", "deleted")
                            
                            # Get last used info
                            last_used = get_venv_last_used_date(venv_path)
                            if last_used:
                                text_widget.insert("end", "Last Used: ", "subheading")
                                text_widget.insert("end", f"{last_used}\n", "timestamp")
                            
                        else:  # association
                            script_path = entry.get("script_path", "")
                            script_dir = entry.get("script_dir", "")
                            script_name = entry.get("script_name", "")
                            venv_path = entry.get("venv_path", "")
                            with_systray = entry.get("with_systray", False)
                            with_startup = entry.get("with_startup", False)
                            
                            text_widget.insert("end", "Script Name: ", "subheading")
                            text_widget.insert("end", f"{script_name}\n", "normal")
                            
                            text_widget.insert("end", "Script Path: ", "subheading")
                            text_widget.insert("end", f"{script_path}\n", "path")
                            
                            text_widget.insert("end", "Script Directory: ", "subheading")
                            text_widget.insert("end", f"{script_dir}\n", "path")
                            
                            text_widget.insert("end", "Virtual Environment Path: ", "subheading")
                            text_widget.insert("end", f"{venv_path}\n", "path")
                            
                            text_widget.insert("end", "System Tray Support: ", "subheading")
                            text_widget.insert("end", f"{with_systray}\n", "normal")
                            
                            text_widget.insert("end", "Startup at Login: ", "subheading")
                            text_widget.insert("end", f"{with_startup}\n", "normal")
                            
                            # Check if association files exist
                            if os.path.exists(script_dir):
                                association_files = [
                                    ".venv-association",
                                    "run_this.vbs",
                                    "run_this.bat",
                                    "launcher.py",
                                    "launcher_icon.ico"
                                ]
                                
                                existing_files = []
                                for file_name in association_files:
                                    file_path = os.path.join(script_dir, file_name)
                                    if os.path.exists(file_path):
                                        existing_files.append(file_name)
                                
                                if existing_files:
                                    text_widget.insert("end", "Association Files Present: ", "subheading")
                                    text_widget.insert("end", ", ".join(existing_files) + "\n", "normal")
                                else:
                                    # Check if venv was deleted
                                    venv_deleted = False
                                    for venv_entry in data:
                                        if venv_entry.get("type") == "venv" and venv_entry.get("venv_path") == venv_path and venv_entry.get("date_deleted"):
                                            venv_deleted = True
                                            break
                                    if date_deleted:
                                        text_widget.insert("end", "No Association Files Present (Deleted through app)\n", "deleted")
                                    elif venv_deleted:
                                        text_widget.insert("end", "No Association Files Present (Deleted due to venv deletion)\n", "deleted")
                                    else:
                                        text_widget.insert("end", "No Association Files Present (Manual Delete Detected Outside of App)\n", "deleted")
                            else:
                                if date_deleted:
                                    text_widget.insert("end", "Script Directory No Longer Exists (Deleted through app)\n", "deleted")
                                else:
                                    text_widget.insert("end", "Script Directory No Longer Exists (Manual Delete Detected Outside of App)\n", "deleted")
                        
                        # Add separator
                        text_widget.insert("end", "\n" + "-" * 80 + "\n\n", "separator")
                        
                except json.JSONDecodeError:
                    # If JSON is invalid, show raw content
                    f.seek(0)
                    text_widget.insert("1.0", f.read())
            
            # Make text widget read-only
            text_widget.configure(state="disabled")
            
            # Add a close button at the bottom
            close_btn = ctk.CTkButton(
                viewer, 
                text="Close", 
                width=100, 
                height=30,
                fg_color="#e74c3c", 
                hover_color="#c0392b", 
                command=viewer.destroy
            )
            close_btn.pack(pady=(0, 10))
            
        except Exception as e:
            self.show_error(f"Error viewing history file: {str(e)}")
    
    def center_dialog(self, dialog, parent=None):
        """Safely center a dialog on its parent window or on screen"""
        if not dialog.winfo_exists():
            return  # Dialog was destroyed, exit early
            
        dialog.update_idletasks()  # Ensure geometry information is up to date
        
        if parent is None or not parent.winfo_exists():
            # Center on screen if parent doesn't exist
            screen_width = dialog.winfo_screenwidth()
            screen_height = dialog.winfo_screenheight()
            dialog_width = dialog.winfo_width()
            dialog_height = dialog.winfo_height()
            
            x = max(0, (screen_width - dialog_width) // 2)
            y = max(0, (screen_height - dialog_height) // 2)
            
            dialog.geometry(f"+{x}+{y}")
            return
        
        # Center on parent if it exists
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_w = parent.winfo_width()
        parent_h = parent.winfo_height()
        dialog_w = dialog.winfo_width()
        dialog_h = dialog.winfo_height()
        
        # Calculate position (prevent negative coordinates)
        x = max(0, parent_x + (parent_w // 2) - (dialog_w // 2))
        y = max(0, parent_y + (parent_h // 2) - (dialog_h // 2))
        
        dialog.geometry(f"+{x}+{y}")

    def export_requirements(self):
        file_path = filedialog.asksaveasfilename(title="Export requirements.txt", defaultextension=".txt", filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if not file_path:
            return
        pkgs = [self.req_listbox.get(idx) for idx in range(self.req_listbox.size())]
        with open(file_path, "w") as f:
            f.write("\n".join(pkgs))
        self.show_success(f"Exported requirements to {file_path}")

    def create_association(self, script_path, venv_path, add_systray=True, add_startup=False):
        """Create an association between a script and a venv, returning a result dictionary"""
        if not script_path or not os.path.exists(script_path):
            return {"success": False, "error": "Invalid script path"}
            
        if not venv_path or not os.path.exists(venv_path):
            return {"success": False, "error": "Invalid venv path"}
            
        # Get script directory and filename
        script_dir = os.path.dirname(script_path)
        script_name = os.path.basename(script_path)
        script_basename = os.path.splitext(script_name)[0]
        
        # Get relative venv path from script directory
        try:
            venv_rel_path = os.path.relpath(venv_path, script_dir)
        except ValueError:
            # If on different drives, use absolute path
            venv_rel_path = venv_path
        
        files_created = []
        
        # Create the association file
        assoc_file_path = os.path.join(script_dir, ".venv-association")
        try:
            with open(assoc_file_path, "w") as f:
                f.write(f"venv_path={venv_rel_path}\n")
                f.write(f"main_script={script_name}\n")
            files_created.append(".venv-association")
        except Exception as e:
            return {"success": False, "error": f"Error creating association file: {str(e)}"}
        
        # Create system tray launcher if option is selected
        if add_systray:
            launcher_path = os.path.join(script_dir, "launcher.py")
            
            # Validate first
            if not script_name or not os.path.exists(os.path.join(script_dir, script_name)):
                return {"success": False, "error": f"Script not found: {script_name}"}

            if not venv_rel_path:
                return {"success": False, "error": "Venv path is missing"}

            # Clean launcher code using f-strings (NO concat)
            launcher_content = f'''import sys
import os
import subprocess
from PySide6.QtWidgets import QApplication, QSystemTrayIcon, QMenu
from PySide6.QtGui import QIcon, QAction, QPixmap, QPainter, QColor, QBrush

# Debug mode flag - set to True to see console output and use python.exe
DEBUG_MODE = False

# Initialize QApplication first
app = QApplication(sys.argv)
app.setQuitOnLastWindowClosed(False)

# Load icon or create fallback AFTER QApplication is created
icon_path = os.path.join(os.path.dirname(__file__), "launcher_icon.ico")
def create_fallback_icon():
    pixmap = QPixmap(64, 64)
    pixmap.fill(QColor(0, 0, 0, 0))
    painter = QPainter(pixmap)
    painter.setBrush(QBrush(QColor(0, 97, 255)))
    painter.drawRect(0, 0, 64, 64)
    painter.setBrush(QBrush(QColor(255, 255, 255)))
    painter.drawRect(16, 16, 32, 32)
    painter.end()
    return QIcon(pixmap)

if os.path.exists(icon_path):
    icon = QIcon(icon_path)
else:
    icon = create_fallback_icon()

tray = QSystemTrayIcon(icon)
menu = QMenu()

# Read script name and venv path from .venv-association
def get_config():
    config = {{}}
    assoc_file = os.path.join(os.path.dirname(__file__), ".venv-association")
    if os.path.exists(assoc_file):
        with open(assoc_file, "r") as f:
            for line in f:
                if "=" in line:
                    key, value = line.strip().split("=", 1)
                    config[key] = value
    return config

config = get_config()
SCRIPT_NAME = config.get("main_script", "main.py")
venv_path = config.get("venv_path", "venv")

def find_python():
    """Find an appropriate Python executable to use based on DEBUG_MODE"""
    script_dir = os.path.dirname(__file__)
    full_venv_path = os.path.join(script_dir, venv_path)
    
    # Choose python executable based on debug mode
    python_exe = "python.exe" if DEBUG_MODE else "pythonw.exe"
    
    # Check multiple possible Python executable locations
    possible_paths = [
        os.path.join(full_venv_path, "Scripts", python_exe),
        os.path.join(full_venv_path, "Scripts", "python.exe"),  # Always check python.exe as fallback
        os.path.join(full_venv_path, "bin", "python"),
        os.path.join(r"C:\Program Files\Python312", python_exe),
        os.path.join(r"C:\Program Files\Python311", python_exe),
        os.path.join(r"C:\Program Files\Python310", python_exe),
        os.path.join(r"C:\Program Files\Python39", python_exe),
        os.path.join(r"C:\Python", python_exe)
    ]
    
    for path in possible_paths:
        if os.path.exists(path):
            return path
            
    # No suitable Python found
    return None

VENV_PYTHON = find_python()

if not VENV_PYTHON:
    # If running in interpreter, print error
    print("[ERROR] No suitable Python interpreter found. Exiting.")
    sys.exit(1)

# Example Exit action
def kill_app():
    if process.poll() is None:
        process.terminate()
    tray.hide()
    app.quit()

exit_action = QAction("Exit App")
exit_action.triggered.connect(kill_app)
menu.addAction(exit_action)
tray.setContextMenu(menu)
tray.setToolTip(f"Launcher for {{SCRIPT_NAME}}")
tray.show()

if not QSystemTrayIcon.isSystemTrayAvailable() or not tray.isVisible():
    print("[ERROR] System tray not available or icon not visible. Exiting launcher.")
    sys.exit(2)
else:
    print("[DEBUG] Tray icon is visible.")

# Verify script exists before launching
script_path = os.path.join(os.path.dirname(__file__), SCRIPT_NAME)
if not os.path.exists(script_path):
    print(f"[ERROR] Script file not found: {{script_path}}")
    sys.exit(3)

try:
    # Prepare command
    cmd = [VENV_PYTHON, SCRIPT_NAME]
    
    # Print command in debug mode
    if DEBUG_MODE:
        print(f"[DEBUG] Running command: {{' '.join(cmd)}}")
    
    # Launch process with or without console based on debug mode
    if DEBUG_MODE:
        process = subprocess.Popen(
            cmd,
            cwd=os.path.dirname(__file__)
        )
    else:
        process = subprocess.Popen(
            cmd,
            cwd=os.path.dirname(__file__),
            creationflags=subprocess.CREATE_NO_WINDOW
        )
    
    print(f"[DEBUG] Launched process PID: {{process.pid}}")
except Exception as e:
    print(f"[ERROR] Failed to launch script: {{str(e)}}")

sys.exit(app.exec())
'''

            # Now write cleanly
            try:
                with open(launcher_path, "w") as f:
                    f.write(launcher_content)
                files_created.append("launcher.py")
            except Exception as e:
                return {"success": False, "error": f"Error creating launcher script: {str(e)}"}
        
        # Create VBS launcher
        try:
            vbs_file_path = os.path.join(script_dir, "run_this.vbs")
            with open(vbs_file_path, "w") as f:
                vbs_content = (
                    'Set WshShell = CreateObject("WScript.Shell")\n'
                    'scriptDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\\"))\n'
                    'Function FindPython()\n'
                    '    Dim paths\n'
                    '    Dim fso\n'
                    '    Dim pythonExe\n'
                    '    Set fso = CreateObject("Scripting.FileSystemObject")\n'
                    '    \n'
                    '    \' First try the relative venv location\n'
                    '    pythonExe = scriptDir & "' + venv_rel_path + '\\Scripts\\pythonw.exe"\n'
                    '    If fso.FileExists(pythonExe) Then\n'
                    '        FindPython = pythonExe\n'
                    '        Exit Function\n'
                    '    End If\n'
                    '    \n'
                    '    \' Try with python.exe instead of pythonw.exe\n'
                    '    pythonExe = scriptDir & "' + venv_rel_path + '\\Scripts\\python.exe"\n'
                    '    If fso.FileExists(pythonExe) Then\n'
                    '        FindPython = pythonExe\n'
                    '        Exit Function\n'
                    '    End If\n'
                    '    \n'
                    '    \' Try common paths\n'
                    '    paths = Array("\\Scripts\\pythonw.exe", "\\Scripts\\python.exe", "\\bin\\python")\n'
                    '    \n'
                    '    For Each path In paths\n'
                    '        pythonExe = scriptDir & "' + venv_rel_path + '" & path\n'
                    '        If fso.FileExists(pythonExe) Then\n'
                    '            FindPython = pythonExe\n'
                    '            Exit Function\n'
                    '        End If\n'
                    '    Next\n'
                    '    \n'
                    '    \' Try system Python locations\n'
                    '    paths = Array("C:\\Program Files\\Python312\\pythonw.exe", "C:\\Program Files\\Python311\\pythonw.exe", "C:\\Program Files\\Python310\\pythonw.exe", "C:\\Program Files\\Python39\\pythonw.exe", "C:\\Python\\pythonw.exe")\n'
                    '    \n'
                    '    For Each path In paths\n'
                    '        If fso.FileExists(path) Then\n'
                    '            FindPython = path\n'
                    '            Exit Function\n'
                    '        End If\n'
                    '    Next\n'
                    '    \n'
                    '    \' Not found\n'
                    '    FindPython = ""\n'
                    'End Function\n'
                )
                
                if add_systray:
                    vbs_content += (
                        'pythonExe = FindPython()\n'
                        'If pythonExe = "" Then\n'
                        '    MsgBox "Python Launcher is sorry to say ..." & vbCrLf & vbCrLf & "No Python found in venv path or system locations", vbExclamation, "Python Launcher Error"\n'
                        '    WScript.Quit\n'
                        'End If\n'
                        'cmd = Chr(34) & pythonExe & Chr(34) & " " & Chr(34) & scriptDir & "launcher.py" & Chr(34)\n'
                    )
                else:
                    vbs_content += (
                        'pythonExe = FindPython()\n'
                        'If pythonExe = "" Then\n'
                        '    MsgBox "Python Launcher is sorry to say ..." & vbCrLf & vbCrLf & "No Python found in venv path or system locations", vbExclamation, "Python Launcher Error"\n'
                        '    WScript.Quit\n'
                        'End If\n'
                        'cmd = Chr(34) & pythonExe & Chr(34) & " " & Chr(34) & scriptDir & "' + script_name + '" & Chr(34)\n'
                    )
                    
                vbs_content += 'WshShell.Run cmd, 0, False\n'
                
                # Add startup registry entry if requested
                if add_startup:
                    vbs_content += '\n' + "' Add to Windows startup\n"
                    vbs_content += 'strRegPath = "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run"\n'
                    vbs_content += f'strRegValue = "{script_basename}"\n'
                    vbs_content += 'WshShell.RegWrite strRegPath & "\\\\" & strRegValue, WScript.ScriptFullName, "REG_SZ"\n'
                    
                f.write(vbs_content)
            files_created.append("run_this.vbs")
        except Exception as e:
            return {"success": False, "error": f"Error creating VBS launcher: {str(e)}"}
            
        # Create BAT launcher with more robust Python detection
        bat_file_path = os.path.join(script_dir, "run_this.bat")
        with open(bat_file_path, "w") as f:
            f.write("@echo off\n")
            f.write("cd /d \"%~dp0\"\n\n")
            f.write(":: Try to find Python in venv\n")
            f.write(f"set PYTHON_EXE={venv_rel_path}\\Scripts\\python.exe\n")
            f.write("if exist %PYTHON_EXE% goto RUN\n\n")
            f.write(f"set PYTHON_EXE={venv_rel_path}\\Scripts\\pythonw.exe\n")
            f.write("if exist %PYTHON_EXE% goto RUN\n\n")
            f.write(f"set PYTHON_EXE={venv_rel_path}\\bin\\python\n")
            f.write("if exist %PYTHON_EXE% goto RUN\n\n")
            f.write(":: Try system Python locations\n")
            f.write("set PYTHON_EXE=C:\\Program Files\\Python312\\python.exe\n")
            f.write("if exist %PYTHON_EXE% goto RUN\n")
            f.write("set PYTHON_EXE=C:\\Program Files\\Python311\\python.exe\n")
            f.write("if exist %PYTHON_EXE% goto RUN\n")
            f.write("set PYTHON_EXE=C:\\Program Files\\Python310\\python.exe\n")
            f.write("if exist %PYTHON_EXE% goto RUN\n")
            f.write("set PYTHON_EXE=C:\\Program Files\\Python39\\python.exe\n")
            f.write("if exist %PYTHON_EXE% goto RUN\n")
            f.write("set PYTHON_EXE=C:\\Python\\python.exe\n")
            f.write("if exist %PYTHON_EXE% goto RUN\n\n")
            f.write("echo Python Launcher is sorry to say ...\n")
            f.write("echo No Python found in venv path or system locations\n")
            f.write("pause\n")
            f.write("exit /b 1\n\n")
            f.write(":RUN\n")
            if add_systray:
                f.write("\"%PYTHON_EXE%\" launcher.py\n")
            else:
                f.write(f"\"%PYTHON_EXE%\" {script_name}\n")
            f.write("pause\n")
        files_created.append("run_this.bat")
        
        if add_systray:
            files_created.append("launcher_icon.ico")
        
        # Log the association in history
        self.log_script_association(script_path, venv_path, add_systray, add_startup)
        
        return {
            "success": True, 
            "script_dir": script_dir,
            "files_created": files_created,
            "add_systray": add_systray,
            "add_startup": add_startup
        }

    def log_script_association(self, script_path, venv_path, add_systray=True, add_startup=False):
        """Log script association to the history file"""
        script_dir = os.path.dirname(script_path)
        script_name = os.path.basename(script_path)
        
        # Get script modification time
        script_mod_time = datetime.datetime.fromtimestamp(os.path.getmtime(script_path)).strftime("%Y-%m-%d %H:%M:%S")
        
        entry = {
            "type": "association",
            "date": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "script_path": script_path,
            "script_dir": script_dir,
            "script_name": script_name,
            "venv_path": venv_path,
            "script_last_modified": script_mod_time,
            "with_systray": add_systray,
            "with_startup": add_startup
        }
        
        try:
            if os.path.exists(ENV_HISTORY_FILE):
                with open(ENV_HISTORY_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
            else:
                data = []
            data.append(entry)
            with open(ENV_HISTORY_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            print(f"Error logging script association: {e}")

    def delete_venv_from_history(self, idx, venv_path, dialog, refresh_history):
        """Delete a venv from the file system and mark it as deleted in history."""
        try:
            # Extract script root directory from venv path
            script_root = os.path.dirname(venv_path) if os.path.exists(venv_path) else None
            
            if script_root:
                # Try to delete the associated files
                self.delete_association_files(script_root)
            
            # Delete venv directory
            if os.path.exists(venv_path):
                try:
                    import shutil
                    # First verify it's a venv directory
                    python_exe = os.path.join(venv_path, "Scripts", "python.exe") if sys.platform == "win32" else os.path.join(venv_path, "bin", "python")
                    if os.path.exists(python_exe):
                        shutil.rmtree(venv_path)
                        # Mark as deleted in history (don't remove the entry)
                        self.log_venv_deletion(venv_path)
                        self.show_success(f"Successfully deleted virtual environment:\n{venv_path}")
                    else:
                        # Safety check to avoid deleting non-venv directories
                        raise ValueError("The selected directory does not appear to be a valid Python virtual environment")
                except Exception as e:
                    message_title = "Error"
                    message_detail = str(e)
                    message_main = "Could not delete virtual environment directory."
                    confirm = ConfirmationDialog(
                        self, message_title, message_main, message_detail,
                        "Continue Anyway", "Cancel", icon_path=self.icon_path
                    )
                    
                    # Create a reference to the dialog for the callback
                    confirm_ref = {"dialog": confirm}
                    
                    # Set action handlers - mark as deleted if user wants to continue anyway
                    confirm.ok_action = lambda: (self.log_venv_deletion(venv_path), 
                                               confirm_ref["dialog"].destroy(), 
                                               refresh_history())
                    confirm.cancel_action = lambda: confirm.destroy()
                    
                    if self.winfo_exists() and confirm.winfo_exists():
                        self.center_dialog(confirm, parent=dialog)
                    return
            else:
                # If the venv doesn't exist, just mark it as deleted
                self.log_venv_deletion(venv_path)
                self.show_success(f"The virtual environment no longer exists. Updated history.")
            
            # Always refresh after operation
            refresh_history()
            
        except Exception as e:
            message_title = "Error"
            message_detail = str(e)
            message_main = "Could not delete virtual environment."
            confirm = ConfirmationDialog(
                self, message_title, message_main, message_detail,
                "OK", "Cancel", icon_path=self.icon_path
            )
            
            # Set action handlers to ensure the dialog is destroyed
            confirm.ok_action = lambda: (confirm.destroy(), refresh_history())
            confirm.cancel_action = lambda: confirm.destroy()
            
            if self.winfo_exists() and confirm.winfo_exists():
                self.center_dialog(confirm, parent=dialog)

    def browse_remove_venv(self, dialog, refresh_history):
        """Browse for a venv to remove."""
        # Get the venv directory to remove
        venv_dir = filedialog.askdirectory(title="Select Virtual Environment to Remove")
        if not venv_dir:
            return
        
        # Check if it's a valid venv
        python_exe = os.path.join(venv_dir, "Scripts", "python.exe") if sys.platform == "win32" else os.path.join(venv_dir, "bin", "python")
        is_venv = os.path.exists(python_exe)
        
        # Extract script root directory from venv path
        script_root = os.path.dirname(venv_dir) if os.path.exists(venv_dir) else None
        
        if script_root:
            # Try to delete the associated files
            self.delete_association_files(script_root)
        
        if not is_venv:
            self.show_error(f"The selected directory does not appear to be a valid Python virtual environment:\n{venv_dir}")
            return
        
        # Ask for confirmation
        message_main = "Are you sure you want to delete this virtual environment?"
        message_detail = f"This will permanently remove the directory:\n{venv_dir}\n\nThis action cannot be undone."
        
        confirm = ConfirmationDialog(
            self, "Confirm Deletion", message_main, message_detail,
            "Delete", "Cancel", icon_path=self.icon_path
        )
        
        # Store reference to the dialog for the callbacks
        confirm_ref = {"dialog": confirm}
        
        def on_ok():
            try:
                import shutil
                shutil.rmtree(venv_dir)
                
                # Log the deletion - only mark as deleted, don't remove from history
                self.log_venv_deletion(venv_dir)
                
                # Update history and refresh
                refresh_history()
                
                # Show success message
                self.show_success(f"Successfully deleted virtual environment:\n{venv_dir}")
            except Exception as e:
                self.show_error(f"Error deleting virtual environment:\n{str(e)}")
            finally:
                if confirm_ref["dialog"].winfo_exists():
                    confirm_ref["dialog"].destroy()
        
        def on_cancel():
            if confirm_ref["dialog"].winfo_exists():
                confirm_ref["dialog"].destroy()
        
        confirm.ok_action = on_ok
        confirm.cancel_action = on_cancel
        
        if self.winfo_exists() and confirm.winfo_exists():
            self.center_dialog(confirm, parent=dialog)

    def show_custom_success(self, message, height=200):
        """Show a success message with custom height"""
        dialog = ctk.CTkToplevel(self)
        dialog.title("Success")
        dialog.geometry(f"400x{height}")  # Custom height
        dialog.transient(self)
        dialog.grab_set()
        
        # Safe focus approach
        if self.winfo_exists():
            dialog.lift()
            if dialog.winfo_exists():
                dialog.focus_set()
        
        self.center_dialog(dialog)
        
        ctk.CTkLabel(dialog, text=message, wraplength=350).pack(pady=20, padx=20)
        
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(pady=10)
        ctk.CTkButton(btn_frame, text="OK", command=dialog.destroy).pack(padx=10)

    def force_delete_entry(self, idx, dialog, refresh_history):
        """Force delete an entry from history without checking for resource existence"""
        try:
            if os.path.exists(ENV_HISTORY_FILE):
                with open(ENV_HISTORY_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                if 0 <= idx < len(data):
                    # Get the entry type for the success message
                    entry_type = data[idx].get("type", "venv")
                    
                    # Remove the entry
                    del data[idx]
                    
                    # Save the modified data
                    with open(ENV_HISTORY_FILE, 'w', encoding='utf-8') as f:
                        json.dump(data, f, indent=2)
                    
                    # Show success message
                    if entry_type == "venv":
                        self.show_success(f"Successfully removed virtual environment entry from the Environment History list.")
                    else:
                        self.show_success(f"Successfully removed script association entry from the Environment History list.")
                    
                    # Refresh the display
                    refresh_history()
                else:
                    self.show_error(f"Invalid index for history entry: {idx}")
            else:
                self.show_error("History file not found.")
        except Exception as e:
            self.show_error(f"Error updating history: {str(e)}")
            refresh_history()  # Refresh anyway to maintain UI consistency

    def add_packages_to_existing_venv(self):
        """Add packages to an existing virtual environment"""
        # Browse for existing venv
        venv_dir = filedialog.askdirectory(title="Select Existing Virtual Environment")
        if not venv_dir:
            return
            
        # Verify it's a valid venv
        python_exe = os.path.join(venv_dir, "Scripts", "python.exe") if sys.platform == "win32" else os.path.join(venv_dir, "bin", "python")
        pip_exe = os.path.join(venv_dir, "Scripts", "pip.exe") if sys.platform == "win32" else os.path.join(venv_dir, "bin", "pip")
        
        if not os.path.exists(python_exe) or not os.path.exists(pip_exe):
            self.show_error(f"The selected directory does not appear to be a valid Python virtual environment:\n{venv_dir}")
            return
            
        # Clear current requirements list
        self.req_listbox.delete(0, tk.END)
        
        # Try to read existing packages from the venv
        try:
            self.show_spinner("Reading installed packages...")
            cmd = [pip_exe, "freeze"]
            result = subprocess.run(cmd, capture_output=True, text=True, check=True)
            packages = [line.strip() for line in result.stdout.splitlines() if line.strip()]
            
            # Filter out packages with version specifiers and clean up
            cleaned_packages = []
            for pkg in packages:
                if "==" in pkg:
                    pkg_name = pkg.split("==")[0].strip()
                    cleaned_packages.append(pkg_name)
                elif pkg and not pkg.startswith("-"):
                    cleaned_packages.append(pkg)
                    
            # Populate the requirements listbox
            for pkg in cleaned_packages:
                self.req_listbox.insert(tk.END, pkg)
                
            self.hide_spinner()
            
            # Store the venv path for later
            self.selected_venv_path = venv_dir
            
            # Update location entry to show the selected venv
            self.location_entry.delete(0, tk.END)
            self.location_entry.insert(0, os.path.dirname(venv_dir))
            
            # Change Create button text to "Update Venv"
            self.create_btn.configure(text="Update Venv", state="normal")
            
            # Show confirmation dialog
            self.show_success(
                f"Loaded {len(cleaned_packages)} packages from the selected virtual environment.\n\n"
                "You can now add or remove packages, and click 'Update Venv' to apply the changes."
            )
            
        except Exception as e:
            self.hide_spinner()
            self.show_error(f"Error reading packages from virtual environment: {str(e)}")
    
    def update_existing_venv(self, venv_path):
        """Update an existing virtual environment with the current requirements list"""
        if not os.path.exists(venv_path):
            self.show_error(f"The virtual environment no longer exists: {venv_path}")
            # Reset the state
            self.selected_venv_path = None
            self.create_btn.configure(text="Create venv")
            return
            
        # Get requirements from listbox, filter stdlib
        requirements = [pkg for pkg in self.req_listbox.get(0, tk.END) if not is_stdlib_module(pkg)]
        skipped = [pkg for pkg in self.req_listbox.get(0, tk.END) if is_stdlib_module(pkg)]
        
        if not requirements:
            self.show_error("No valid packages selected. Please add packages to install.")
            return
        
        def update_task():
            self.show_spinner("Updating virtual environment...")
            try:
                # Install requirements using pip
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
                
                # Show installation summary
                self.show_install_summary(success, failed, skipped)
                
                # Reset state
                self.selected_venv_path = None
                self.create_btn.configure(text="Create venv")
                
                # Open the venv directory
                try:
                    if platform.system() == "Windows":
                        os.startfile(venv_path)
                    elif platform.system() == "Darwin":
                        subprocess.Popen(["open", venv_path])
                    else:
                        subprocess.Popen(["xdg-open", venv_path])
                except Exception as e:
                    print(f"Error opening venv directory: {e}")
                
                # Clear UI
                self.req_listbox.delete(0, tk.END)
                self.location_entry.delete(0, tk.END)
                
            except Exception as e:
                self.hide_spinner()
                self.show_error(f"Error updating virtual environment: {str(e)}")
                # Reset state
                self.selected_venv_path = None
                self.create_btn.configure(text="Create venv")
        
        threading.Thread(target=update_task, daemon=True).start()

def get_venv_last_used_date(venv_path):
    """Standalone version of the venv last used date function"""
    if not venv_path or not os.path.exists(venv_path):
        return ""
        
    # Path to activate script
    activate_path = os.path.join(venv_path, "Scripts" if sys.platform == "win32" else "bin", "activate")
    
    # Check if activate script exists
    if os.path.exists(activate_path):
        try:
            # Get the modification time
            mod_time = os.path.getmtime(activate_path)
            # Convert to datetime and format
            return datetime.datetime.fromtimestamp(mod_time).strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            pass
            
    return ""

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

class ConfirmationDialog(ctk.CTkToplevel):
    """Dialog for confirming actions with two options."""
    def __init__(self, master, title, message_main, message_detail, option1_text, option2_text, icon_path=None):
        super().__init__(master)
        self.title(title)
        self.geometry("500x300")
        self.transient(master)
        self.grab_set()
        self.lift()
        self.focus_set()
        
        # Set icon if available
        if icon_path and os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass
        
        # Main message (bold)
        ctk.CTkLabel(self, text=message_main, font=("Arial", 12, "bold")).pack(pady=(20, 10), padx=20)
        
        # Detail message
        ctk.CTkLabel(self, text=message_detail, wraplength=460).pack(pady=(0, 20), padx=20)
        
        # Button frame with fixed height
        btn_frame = ctk.CTkFrame(self, fg_color="transparent", height=50)
        btn_frame.pack(pady=10, fill="x", padx=20)
        btn_frame.pack_propagate(False)  # Prevent frame from shrinking
        
        # Option 1 button (default action) - LEFT
        self.option1_btn = ctk.CTkButton(
            btn_frame, 
            text=option1_text, 
            width=140, 
            height=32,
            fg_color="#1A73E8", 
            hover_color="#1557B0", 
            command=self._on_option1
        )
        self.option1_btn.pack(side="left", padx=(40, 0))
        
        # Option 2 button (cancel) - RIGHT
        self.option2_btn = ctk.CTkButton(
            btn_frame, 
            text=option2_text, 
            width=140, 
            height=32,
            fg_color="#e74c3c",  # Red cancel button for better visibility
            hover_color="#c0392b", 
            command=self._on_option2
        )
        self.option2_btn.pack(side="right", padx=(0, 40))
        
        # Action callbacks, to be set by caller
        self.ok_action = lambda: self.destroy()
        self.cancel_action = lambda: self.destroy()
        
    def _on_option1(self):
        """Execute the primary action."""
        self.ok_action()
        
    def _on_option2(self):
        """Execute the secondary action."""
        self.cancel_action()
        
    def _on_cancel(self):
        """Handle dialog cancellation."""
        self.cancel_action()
        self.destroy()

    def view_history_file(self, parent_dialog):
        """Open a dialog to view the history file contents"""
        try:
            # Check if history file exists
            if not os.path.exists(ENV_HISTORY_FILE):
                self.show_error("History file does not exist yet.")
                return
                
            # Create a new dialog to display the history content
            viewer = ctk.CTkToplevel(self)
            viewer.title("Environment History File")
            viewer.geometry("800x600")
            viewer.transient(parent_dialog)
            viewer.grab_set()
            viewer.lift()
            
            # Center the dialog
            self.center_dialog(viewer, parent=parent_dialog)
            
            # Create a frame for the text widget
            frame = ctk.CTkFrame(viewer)
            frame.pack(fill="both", expand=True, padx=10, pady=10)
            
            # Create a text widget with scrollbar
            text_widget = tk.Text(
                frame, 
                wrap="none", 
                bg="#232830", 
                fg="white",
                font=("Consolas", 11),
                padx=10,
                pady=10
            )
            
            # Add scrollbars
            y_scrollbar = tk.Scrollbar(frame, orient="vertical", command=text_widget.yview)
            x_scrollbar = tk.Scrollbar(frame, orient="horizontal", command=text_widget.xview)
            text_widget.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
            
            # Pack scrollbars and text widget
            y_scrollbar.pack(side="right", fill="y")
            x_scrollbar.pack(side="bottom", fill="x")
            text_widget.pack(side="left", fill="both", expand=True)
            
            # Load and insert the history file contents
            with open(ENV_HISTORY_FILE, "r", encoding="utf-8") as f:
                # Pretty format the JSON for better readability
                import json
                try:
                    data = json.load(f)
                    formatted_json = json.dumps(data, indent=4)
                    text_widget.insert("1.0", formatted_json)
                except json.JSONDecodeError:
                    # If JSON is invalid, show raw content
                    f.seek(0)
                    text_widget.insert("1.0", f.read())
            
            # Make text widget read-only
            text_widget.configure(state="disabled")
            
            # Add a close button at the bottom
            close_btn = ctk.CTkButton(
                viewer, 
                text="Close", 
                width=100, 
                height=30,
                fg_color="#e74c3c", 
                hover_color="#c0392b", 
                command=viewer.destroy
            )
            close_btn.pack(pady=(0, 10))
            
        except Exception as e:
            self.show_error(f"Error viewing history file: {str(e)}")

    def show_info(self, message):
        """Show an informational message with safe window handling"""
        dialog = ctk.CTkToplevel(self)
        dialog.title("Information")
        dialog.geometry("400x150")
        dialog.transient(self)
        dialog.grab_set()
        
        # Safe focus approach
        if self.winfo_exists():
            dialog.lift()
            if dialog.winfo_exists():
                dialog.focus_set()
                
        # Center the dialog only if it still exists
        if dialog.winfo_exists():
            self.center_dialog(dialog)
            
        ctk.CTkLabel(dialog, text=message, wraplength=350).pack(pady=20, padx=20)
        ctk.CTkButton(dialog, text="OK", command=dialog.destroy).pack(pady=10)

    def copy_to_clipboard(self, text):
        """Copy text to clipboard"""
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        root.clipboard_clear()
        root.clipboard_append(text)
        root.update()  # now it stays on the clipboard after the window is closed
        root.destroy()