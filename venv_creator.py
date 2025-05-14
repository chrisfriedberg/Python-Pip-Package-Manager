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
        """Associate a Python script with a virtual environment and create launcher files"""
        # Dialog for selecting script and venv
        dialog = ctk.CTkToplevel(self)
        dialog.title("Associate Script with Virtual Environment")
        dialog.geometry("600x480")
        dialog.transient(self)
        dialog.grab_set()
        if self.winfo_exists() and dialog.winfo_exists():
            if dialog.winfo_exists():
                dialog.lift()
            if dialog.winfo_exists():
                dialog.focus_set()
            if dialog.winfo_exists():
                self.center_dialog(dialog)
        
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
        
        def create_association():
            script_path = script_entry.get().strip()
            venv_path = venv_entry.get().strip()
            
            if not script_path or not os.path.exists(script_path):
                self.show_error("Please select a valid Python script.")
                return
                
            if not venv_path or not os.path.exists(venv_path):
                self.show_error("Please select a valid virtual environment directory.")
                return
                
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
            
            # Create the association file
            assoc_file_path = os.path.join(script_dir, ".venv-association")
            try:
                with open(assoc_file_path, "w") as f:
                    f.write(f"venv_path={venv_rel_path}\n")
                    f.write(f"main_script={script_name}\n")
            except Exception as e:
                self.show_error(f"Error creating association file: {str(e)}")
                return
            
            # Create system tray launcher if option is selected
            if systray_var.get():
                launcher_path = os.path.join(script_dir, "launcher.py")
                
                # Validate first
                if not script_name or not os.path.exists(os.path.join(script_dir, script_name)):
                    self.show_error(f"Script not found: {script_name}")
                    return

                if not venv_rel_path:
                    self.show_error("Venv path is missing")
                    return

                # Clean launcher code using f-strings (NO concat)
                launcher_content = f'''import sys
import os
import subprocess
from PySide6.QtWidgets import QApplication, QSystemTrayIcon, QMenu
from PySide6.QtGui import QIcon, QAction, QPixmap, QPainter, QColor, QBrush

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

# Read script name from .venv-association
def get_script_name():
    config = {{}}
    assoc_file = os.path.join(os.path.dirname(__file__), ".venv-association")
    if os.path.exists(assoc_file):
        with open(assoc_file, "r") as f:
            for line in f:
                if "=" in line:
                    key, value = line.strip().split("=", 1)
                    config[key] = value
    return config.get("main_script", "main.py")

SCRIPT_NAME = get_script_name()
VENV_PYTHON = os.path.join(os.path.dirname(__file__), 'venv', 'Scripts', 'pythonw.exe')

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

try:
    process = subprocess.Popen(
        [VENV_PYTHON, SCRIPT_NAME],
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
                except Exception as e:
                    self.show_error(f"Error creating launcher script: {str(e)}")
                    return
            
            # Create VBS launcher
            try:
                vbs_file_path = os.path.join(script_dir, "run_this.vbs")
                with open(vbs_file_path, "w") as f:
                    vbs_content = (
                        'Set WshShell = CreateObject("WScript.Shell")\n'
                        'scriptDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\\"))\n'
                    )
                    if systray_var.get():
                        vbs_content += (
                            'cmd = Chr(34) & scriptDir & "venv\\Scripts\\pythonw.exe" & Chr(34) & " " & Chr(34) & scriptDir & "launcher.py" & Chr(34)\n'
                        )
                    else:
                        vbs_content += (
                            f'cmd = Chr(34) & scriptDir & "venv\\Scripts\\pythonw.exe" & Chr(34) & " " & Chr(34) & scriptDir & "{script_name}" & Chr(34)\n'
                        )
                    vbs_content += 'WshShell.Run cmd, 0, False\n'
                    # Add startup registry entry if requested
                    if startup_var.get():
                        vbs_content += '\n' + "' Add to Windows startup\n"
                        vbs_content += 'strRegPath = "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run"\n'
                        vbs_content += f'strRegValue = "{script_basename}"\n'
                        vbs_content += 'WshShell.RegWrite strRegPath & "\\\\" & strRegValue, WScript.ScriptFullName, "REG_SZ"\n'
                    f.write(vbs_content)
            except Exception as e:
                self.show_error(f"Error creating VBS launcher: {str(e)}")
                return
                
            # Create BAT launcher
            bat_file_path = os.path.join(script_dir, "run_this.bat")
            with open(bat_file_path, "w") as f:
                f.write("@echo off\n")
                f.write("cd /d \"%~dp0\"\n")
                f.write("venv\\Scripts\\python.exe launcher.py\n")
                f.write("pause\n")
            
            # Files list for success message
            files_created = [
                "- .venv-association",
                "- run_this.vbs (silent launcher)",
                "- run_this.bat (CLI launcher)"
            ]
            
            if systray_var.get():
                files_created.append("- launcher.py (system tray control)")
                files_created.append("- launcher_icon.ico (automatically generated)")
                
            # Show success message with details
            success_msg = (
                f"Successfully created association files in:\n{script_dir}\n\n"
                "Files created:\n" + "\n".join(files_created) + "\n\n"
            )
            
            if systray_var.get():
                success_msg += "You can now run your script by launching launcher.py\n"
                success_msg += "Right-click the tray icon to access the Exit option."
            else:
                success_msg += "You can now run your script by double-clicking run_this.vbs"
                
            if startup_var.get():
                success_msg += "\nThe application will also start automatically at Windows login."
                
            # Mention PySide6 installation
            if systray_var.get():
                success_msg += "\n\nNote: The tray icon uses PySide6, which will be installed in the venv when needed."
            
            # Show success dialog BEFORE destroying the association dialog
            success_dialog = ctk.CTkToplevel(self)
            success_dialog.title("Success")
            success_dialog.geometry("500x300")
            success_dialog.transient(self)
            success_dialog.grab_set()
            if self.winfo_exists() and success_dialog.winfo_exists():
                if success_dialog.winfo_exists():
                    success_dialog.lift()
                if success_dialog.winfo_exists():
                    success_dialog.focus_set()
                if success_dialog.winfo_exists():
                    self.center_dialog(success_dialog)
            ctk.CTkLabel(success_dialog, text=success_msg, wraplength=480, justify="left").pack(pady=20, padx=20)
            def close_both():
                if success_dialog.winfo_exists():
                    success_dialog.destroy()
                if dialog.winfo_exists():
                    dialog.destroy()
            ctk.CTkButton(success_dialog, text="OK", command=close_both).pack(pady=10)
        
        # Create and Cancel buttons
        create_btn = ctk.CTkButton(
            btn_frame, 
            text="Create Association", 
            width=150, 
            fg_color="#1A73E8", 
            hover_color="#1557B0", 
            command=create_association
        )
        create_btn.pack(side="left", padx=5)
        
        cancel_btn = ctk.CTkButton(
            btn_frame, 
            text="Cancel", 
            width=100, 
            fg_color="#e74c3c", 
            hover_color="#b93a2b", 
            command=dialog.destroy
        )
        cancel_btn.pack(side="right", padx=5)

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
        dialog.geometry("1080x480")  # Increased width from 960 to 1080
        dialog.transient(self)
        dialog.grab_set()
        dialog.lift()
        dialog.focus_set()
        self.center_dialog(dialog)
        
        # Hide deleted venvs checkbox (above table)
        hide_deleted_var = tk.BooleanVar(value=False)
        hide_cb = ctk.CTkCheckBox(dialog, text="Hide deleted venvs", variable=hide_deleted_var)
        hide_cb.pack(anchor="w", padx=20, pady=(10, 0))

        # Table frame
        table_frame = ctk.CTkFrame(dialog, fg_color="#23272E")
        table_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Sorting state
        sort_state = {"column": None, "reverse": False}

        # Header row with clickable headers for sorting
        headers = [
            ("Date Added", 120),
            ("Last Used", 120),  # Based on activate script modification time
            ("Date Removed", 120),
            ("Script Root", 120),
            ("Venv Path", 300),
            ("Actions", 120)  # For buttons
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
            
        def venv_exists(venv_path):
            """Check if the venv directory actually exists"""
            return venv_path and os.path.exists(venv_path) and os.path.isdir(venv_path)

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
                
            # Add dynamic data for sorting and display
            for entry in data:
                # Get last used date from file modification time
                entry["_last_used"] = get_venv_last_used_date(entry.get("venv_path", ""))
                # Check if venv directory actually exists
                entry["_venv_exists"] = venv_exists(entry.get("venv_path", ""))
                
            # Apply sorting if a column is selected
            if sort_state["column"] is not None:
                col_idx = sort_state["column"]
                reverse = sort_state["reverse"]
                
                # Map column index to the appropriate sorting key
                sort_keys = {
                    0: lambda e: e.get("date", ""),  # Date Added
                    1: lambda e: e.get("_last_used", ""),  # Last Used (from file mod time)
                    2: lambda e: e.get("date_deleted", ""),  # Date Removed
                    3: lambda e: e.get("parent", "").lower(),  # Script Root
                    4: lambda e: e.get("venv_path", "").lower(),  # Venv Path
                }
                
                if col_idx in sort_keys:
                    data.sort(key=sort_keys[col_idx], reverse=reverse)
            
            for i, entry in enumerate(data):
                row = i + 1  # header is row 0
                date_added = entry.get("date", "")
                last_used = entry.get("_last_used", "")  # From activate script modification time
                date_removed = entry.get("date_deleted", "")
                script_root = entry.get("parent", "")
                venv_path = entry.get("venv_path", "")
                venv_exists_flag = entry.get("_venv_exists", False)
                
                ctk.CTkLabel(table_frame, text=date_added, width=120, anchor="w").grid(row=row, column=0, padx=5, sticky="w")
                ctk.CTkLabel(table_frame, text=last_used, width=120, anchor="w").grid(row=row, column=1, padx=5, sticky="w")
                ctk.CTkLabel(table_frame, text=date_removed, width=120, anchor="w").grid(row=row, column=2, padx=5, sticky="w")
                ctk.CTkLabel(table_frame, text=script_root, width=120, anchor="w").grid(row=row, column=3, padx=5, sticky="w")
                
                # Show path with appropriate color - red if it doesn't exist
                path_color = "#1A73E8" if venv_exists_flag else "#e74c3c"
                
                def open_folder(path):
                    folder = os.path.dirname(path)
                    if os.path.exists(folder):
                        if platform.system() == "Windows":
                            os.startfile(folder)
                        elif platform.system() == "Darwin":
                            subprocess.Popen(["open", folder])
                        else:
                            subprocess.Popen(["xdg-open", folder])
                    else:
                        self.show_error(f"Folder does not exist: {folder}")
                
                def make_link_callback(p=venv_path):
                    return lambda e: open_folder(p)
                    
                venv_path_label = ctk.CTkLabel(table_frame, text=venv_path, width=300, anchor="w", 
                                            text_color=path_color, cursor="hand2")
                venv_path_label.grid(row=row, column=4, padx=5, sticky="w")
                venv_path_label.bind("<Button-1>", make_link_callback())
                
                # Create action buttons frame
                btn_frame = ctk.CTkFrame(table_frame, fg_color="transparent")
                btn_frame.grid(row=row, column=5, padx=5, sticky="w")
                
                # Delete Entry button
                del_entry_btn = ctk.CTkButton(
                    btn_frame, 
                    text="Delete Entry", 
                    width=58, 
                    height=24,
                    fg_color="#FF8C00", 
                    hover_color="#E57C00",
                    command=lambda idx=i: self.remove_venv_entry(idx, dialog, refresh_history)
                )
                del_entry_btn.pack(side="left", padx=(0, 4))
                
                # Only show the Delete venv button if venv exists and isn't already marked as deleted
                if venv_exists_flag and not date_removed:
                    del_venv_btn = ctk.CTkButton(
                        btn_frame, 
                        text="Delete venv", 
                        width=58, 
                        height=24,
                        fg_color="#e74c3c", 
                        hover_color="#b93a2b", 
                        command=lambda idx=i, path=venv_path: self.delete_venv_from_history(idx, path, dialog, refresh_history)
                    )
                    del_venv_btn.pack(side="left")
                
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
        
        # Button dimensions for right-side buttons
        btn_width = 180
        btn_height = 32
        
        # Close button (bottom right, under browse_remove_btn)
        close_btn = ctk.CTkButton(dialog, text="Close", width=btn_width, height=btn_height, fg_color="#e74c3c", hover_color="#b93a2b", command=dialog.destroy)
        close_btn.place(relx=1.0, rely=1.0, x=-20, y=-20-btn_height-10, anchor="se")
        
        # Browse To Remove venv button (bottom right, blue)
        browse_remove_btn = ctk.CTkButton(dialog, text="Browse To Remove venv", width=btn_width, height=btn_height, fg_color="#1565c0", hover_color="#0d47a1", command=lambda: self.browse_remove_venv(dialog, refresh_history))
        browse_remove_btn.place(relx=1.0, rely=1.0, x=-20, y=-20, anchor="se")
        
        hide_cb.configure(command=refresh_history)
        refresh_history()
        
    def remove_venv_entry(self, idx, dialog, refresh_callback):
        """Completely remove an entry from the history file"""
        try:
            if os.path.exists(VENVS_HISTORY_FILE):
                with open(VENVS_HISTORY_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
            else:
                data = []
                
            if 0 <= idx < len(data):
                # Ask for confirmation
                confirm = ctk.CTkToplevel(dialog)
                confirm.title("Confirm Delete Entry")
                confirm.geometry("420x180")
                confirm.transient(dialog)
                confirm.grab_set()
                confirm.lift()
                confirm.focus_set()
                self.center_dialog(confirm, parent=dialog)
                
                confirm_text = f"Are you sure you want to completely remove this entry from history?\nThis only removes the record, not the actual venv."
                ctk.CTkLabel(confirm, text=confirm_text, wraplength=380).pack(pady=20, padx=20)
                
                btn_frame = ctk.CTkFrame(confirm, fg_color="transparent")
                btn_frame.pack(pady=10)
                
                result = {'ok': False}
                def on_ok():
                    result['ok'] = True
                    confirm.destroy()
                def on_cancel():
                    confirm.destroy()
                    
                ctk.CTkButton(btn_frame, text="Delete Entry", fg_color="#FF8C00", hover_color="#E57C00", command=on_ok).pack(side="left", padx=10)
                ctk.CTkButton(btn_frame, text="Cancel", command=on_cancel).pack(side="right", padx=10)
                confirm.wait_window()
                
                if result['ok']:
                    # Remove the entry
                    entry_name = data[idx].get("parent", "Unknown")
                    del data[idx]
                    
                    # Save the modified data
                    with open(VENVS_HISTORY_FILE, "w", encoding="utf-8") as f:
                        json.dump(data, f, indent=2)
                        
                    # Show success message
                    self.show_success(f"Entry for '{entry_name}' has been removed from history.")
                    
                    # Refresh the display
                    refresh_callback()
                    
        except Exception as e:
            self.show_error(f"Error removing entry: {str(e)}")
            
    def delete_venv_from_history(self, idx, venv_path, dialog, refresh_callback):
        """Delete the actual venv directory and mark it as deleted in history"""
        try:
            # Verify the venv exists before trying to delete it
            if not os.path.exists(venv_path) or not os.path.isdir(venv_path):
                self.show_error(f"Venv directory no longer exists: {venv_path}")
                refresh_callback()  # Refresh to update UI state
                return
                
            if not os.path.basename(venv_path) in ['venv', '.venv']:
                self.show_error("Only directories named 'venv' or '.venv' can be deleted.")
                return
                
            # Check if we're trying to delete the venv that the application is running in
            is_app_venv = False
            app_venv_path = sys.prefix if (hasattr(sys, 'real_prefix') or 
                        (hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix)) else None
            if app_venv_path:
                # Normalize paths for comparison
                venv_dir_norm = os.path.normpath(venv_path)
                app_venv_norm = os.path.normpath(app_venv_path)
                is_app_venv = venv_dir_norm == app_venv_norm
                
            # Confirm deletion
            confirm = ctk.CTkToplevel(dialog)
            confirm.title("Confirm Delete venv")
            confirm.geometry("480x220" if is_app_venv else "420x180")
            confirm.transient(dialog)
            confirm.grab_set()
            confirm.lift()
            confirm.focus_set()
            self.center_dialog(confirm, parent=dialog)
            
            warning_text = f"Are you sure you want to delete this venv?\n{venv_path}"
            if is_app_venv:
                warning_text += "\n\nWARNING: This appears to be the virtual environment that the Package Manager is currently running in. Deleting it will remove all packages needed by this application. You will need to reinstall them manually after deletion."
                
            ctk.CTkLabel(confirm, text=warning_text, wraplength=440, text_color="#e74c3c" if is_app_venv else "white").pack(pady=20, padx=20)
            btn_frame = ctk.CTkFrame(confirm, fg_color="transparent")
            btn_frame.pack(pady=10)
            
            result = {'ok': False}
            def on_ok():
                result['ok'] = True
                confirm.destroy()
            def on_cancel():
                confirm.destroy()
                
            ctk.CTkButton(btn_frame, text="Delete" if not is_app_venv else "Delete Anyway", fg_color="#e74c3c", hover_color="#b93a2b", command=on_ok).pack(side="left", padx=10)
            ctk.CTkButton(btn_frame, text="Cancel", command=on_cancel).pack(side="right", padx=10)
            
            confirm.wait_window()
            if not result['ok']:
                return
                
            # Delete the folder
            import shutil
            try:
                shutil.rmtree(venv_path)
                self.log_venv_deletion(venv_path)
                
                msg = f"Deleted venv: {venv_path}"
                if is_app_venv:
                    msg += "\n\nYou have deleted the virtual environment used by this application. Some features may not work properly. Please reinstall the required packages using:\n\npip install -r requirements.txt"
                    
                self.show_success(msg)
                
                # Update the entry and refresh
                if os.path.exists(VENVS_HISTORY_FILE):
                    with open(VENVS_HISTORY_FILE, "r", encoding="utf-8") as f:
                        data = json.load(f)
                        
                    if 0 <= idx < len(data):
                        data[idx]["date_deleted"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        
                        with open(VENVS_HISTORY_FILE, "w", encoding="utf-8") as f:
                            json.dump(data, f, indent=2)
                
                refresh_callback()
            except Exception as e:
                self.show_error(f"Error deleting venv: {e}")
                refresh_callback()  # Still refresh to update UI
                
        except Exception as e:
            self.show_error(f"Error processing deletion: {str(e)}")
            refresh_callback()
            
    def export_venv_history(self, parent_dialog):
        """Export venv history to a text file"""
        file_path = filedialog.asksaveasfilename(
            parent=parent_dialog,
            title="Export Venv History", 
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
            
        try:
            if os.path.exists(VENVS_HISTORY_FILE):
                with open(VENVS_HISTORY_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
            else:
                data = []
                
            with open(file_path, "w", encoding="utf-8") as f:
                f.write("Virtual Environment History Log\n")
                f.write("============================\n\n")
                
                for i, entry in enumerate(data):
                    date_added = entry.get("date", "N/A")
                    last_used = get_venv_last_used_date(entry.get("venv_path", ""))
                    date_removed = entry.get("date_deleted", "")
                    script_root = entry.get("parent", "")
                    venv_path = entry.get("venv_path", "")
                    exists = "Yes" if os.path.exists(venv_path) else "No"
                    
                    f.write(f"Entry #{i+1}:\n")
                    f.write(f"  Date Added: {date_added}\n")
                    f.write(f"  Last Used: {last_used if last_used else 'Never/Unknown'}\n")
                    f.write(f"  Date Removed: {date_removed if date_removed else 'N/A'}\n")
                    f.write(f"  Script Root: {script_root}\n")
                    f.write(f"  Venv Path: {venv_path}\n")
                    f.write(f"  Still Exists: {exists}\n\n")
                    
            self.show_success(f"Exported venv history to {file_path}")
            
        except Exception as e:
            self.show_error(f"Error exporting history: {str(e)}")
            
    def center_dialog(self, dialog, parent=None):
        """Safely center a dialog on its parent window"""
        if not dialog.winfo_exists():
            return  # Dialog was destroyed, exit early
            
        dialog.update_idletasks()
        if parent is None:
            parent = self
            
        if not parent.winfo_exists():
            # Fallback to screen center if parent no longer exists
            screen_width = dialog.winfo_screenwidth()
            screen_height = dialog.winfo_screenheight()
            x = (screen_width - dialog.winfo_width()) // 2
            y = (screen_height - dialog.winfo_height()) // 2
            dialog.geometry(f"+{x}+{y}")
            return
            
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
        
    def browse_remove_venv(self, dialog, refresh_history):
        from tkinter import filedialog
        venv_dir = filedialog.askdirectory(title="Select venv folder to delete")
        if not venv_dir:
            return
        # Confirm deletion only if the directory is named 'venv' or '.venv'
        if not os.path.basename(venv_dir) in ['venv', '.venv']:
            self.show_error("You can only delete directories named 'venv' or '.venv'.")
            return
            
        # Verify the directory exists
        if not os.path.exists(venv_dir) or not os.path.isdir(venv_dir):
            self.show_error(f"Selected directory does not exist: {venv_dir}")
            return
            
        # Find index in history if it exists
        idx = -1
        venv_path = venv_dir
        
        try:
            if os.path.exists(VENVS_HISTORY_FILE):
                with open(VENVS_HISTORY_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    
                # Find the index of this venv path in the data
                for i, entry in enumerate(data):
                    if entry["venv_path"] == venv_path:
                        idx = i
                        break
        except Exception:
            pass
            
        # Use delete_venv_from_history to handle the deletion
        self.delete_venv_from_history(idx, venv_path, dialog, refresh_history)
        
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