import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys
import importlib.metadata
import importlib.util
import datetime
import subprocess
import threading
import platform
import time
import re
import string
from collections import defaultdict
import ctypes
import json
import configparser
import getpass

# --- Configurations ---
# These are less critical now that we're not scanning system-wide,
# but can still be useful for sub-folders within user-selected directories.
EXCLUDED_SUBFOLDER_NAMES = { # Common folders to skip within a project
    "__pycache__", ".git", ".hg", ".svn", "venv", ".venv", "env", ".env",
    "node_modules", "build", "dist", "target", "docs", "tests", "test" 
}

JUNK_IMPORT_WORDS = {
    "__main__", "__builtin__", "__init__", "self", "cls", "args", "kwargs", "true", "false", "none",
    "if", "else", "elif", "for", "while", "try", "except", "finally", "with", "as", "import", "from", "def", "class",
    "return", "yield", "lambda", "pass", "continue", "break", "global", "nonlocal", "del", "assert",
    "a", "an", "the", "it", "be", "to", "of", "for", "on", "at", "by", "this", "that", "all", "any", "some",
    "main", "util", "utils", "helper", "helpers", "config", "setup", "example", "examples",
    "script", "scripts", "tool", "tools", "lib", "core", "app", "src", "pkg", "module", "package", "project",
    "data", "api", "client", "server", "common", "base", "interface", "model", "models", "view", "views",
    "string", "int", "float", "bool", "list", "dict", "set", "tuple", "object", "function", "method",
    "variable", "constant", "parameter", "argument", "index", "item", "value", "result", "status", "code",
    "name", "file", "path", "root", "dir", "folder", "user", "users", "system", "windows", "linux", "mac",
    "http", "https", "url", "uri", "ip", "port", "host", "json", "xml", "csv", "yaml", "html", "css", "js",
    "date", "time", "datetime", "timestamp", "year", "month", "day", "hour", "minute", "second",
    "foo", "bar", "baz", "qux",
    "_abcoll", "_manylinux", "_pydev_bundle", "_pydev_runfiles", "_pydevd_bundle", "_pytest",
    "_pydevd_bundle_ext", "_pydevd_frame_eval", "_pydevd_frame_eval_ext",
    "_pydevd_sys_monitoring", "_pypy__", "_pypy_wait", "_subprocess", "_typeshed",
    "distutils", "pkg_resources", "py", "python",
    # Adding local module names that were causing false positives
    "documentversionexplorer", "log_workspace_action", "model_loader", "openaiwhisper",
    "venv_creator", "video_frame_snatcher", "win32clipboard", "win32con", "win32file",
    "youtube_caption_fetcher", "youtube_captionfetcher"  # Adding both variations
}
PROTECTED_PACKAGES = {
    "pip", "setuptools", "wheel", "customtkinter", "python-dotenv",
    "certifi", "charset-normalizer", "idna", "requests", "urllib3", 
    "packaging", "pyparsing", "colorama", "python", "py", "openai-whisper"
}
USER_STDLIB_MODULE_LIST = [
    "abc", "argparse", "asyncio", "base64", "collections", "concurrent", "contextlib", "copy", "csv", "datetime",
    "decimal", "difflib", "enum", "fileinput", "fnmatch", "functools", "glob", "gzip", "hashlib", "heapq",
    "hmac", "html", "http", "importlib", "inspect", "io", "itertools", "json", "logging", "math",
    "multiprocessing", "operator", "os", "pathlib", "pickle", "platform", "pprint", "queue", "random",
    "re", "selectors", "shutil", "signal", "socket", "sqlite3", "ssl", "stat", "string", "struct",
    "subprocess", "sys", "tempfile", "threading", "time", "timeit", "types", "typing", "unittest",
    "urllib", "uuid", "warnings", "weakref", "xml", "xmlrpc", "zipfile", "zlib", "configparser", "email",
    "imp", "msvcrt", "winsound", "winreg", "sysconfig", "_thread", "builtins"
]
USER_MODULE_TO_PACKAGE_MAPPING = {
    "PIL": "pillow", "sklearn": "scikit-learn", "cv2": "opencv-python", "bs4": "beautifulsoup4",
    "wx": "wxPython", "tk": "tk", "tkinter": "tkinter", "matplotlib": "matplotlib", "np": "numpy",
    "pd": "pandas", "plt": "matplotlib", "customtkinter": "customtkinter", "ctk": "customtkinter",
    "pyside6": "PySide6", "win32com": "pywin32", "win32api": "pywin32", "win32gui": "pywin32",
    "pythoncom": "pywin32", "pywintypes": "pywin32",
    "sentence_transformers": "sentence-transformers",  # Adding proper mapping
    "openaiwhisper": "openai-whisper",  # Adding proper mapping
    "whisper": "openai-whisper",  # Also map the base module name
    "openai_whisper": "openai-whisper",  # Add underscore variant
    "OpenAIWhisper": "openai-whisper",  # Add CamelCase variant
}
USER_IMPORT_PATTERNS = [
    r'^\s*import\s+([a-zA-Z_][a-zA-Z0-9_.]*)',
    r'^\s*from\s+([a-zA-Z_][a-zA-Z0-9_.]*)\s+import',
    r'^\s*import\s+([a-zA-Z_][a-zA-Z0-9_.]*)\s+as\s+\w+'
]

SETTINGS_INI = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'settings.ini')

DEFAULT_EXCLUDE_DIRS = [
    os.path.join('%USERPROFILE%', 'AppData'),
    os.path.join('%USERPROFILE%', 'Music'),
    os.path.join('%USERPROFILE%', 'Searches'),
    os.path.join('%USERPROFILE%', 'AppData', 'Roaming'),
    os.path.join('%USERPROFILE%', 'Downloads'),
    os.path.join('%USERPROFILE%', 'Favorites'),
]

class RequirementsDoctor(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")
        self.title("Targeted Requirements Doctor")
        self.geometry("800x750")
        self.cancel_event = threading.Event()
        self.user_home_dir = os.path.expanduser('~').lower()
        self.queued_directories = []
        self.always_include = set()
        self.always_uninstall = set()
        self.exclude_dirs = set()
        self.load_settings_ini()

        # --- Main Container Frame ---
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=30, pady=30)
        self.main_frame.grid_columnconfigure(0, weight=1)

        # --- Directory Management Section ---
        dir_section = ctk.CTkFrame(self.main_frame, fg_color=("#23272e", "#23272e"))
        dir_section.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        dir_section.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(dir_section, text="Directory Management", font=("Arial", 18, "bold")).grid(row=0, column=0, sticky="w", padx=10, pady=(10, 0))
        self.dir_list_frame = ctk.CTkScrollableFrame(dir_section, label_text="Selected Directories", height=180)
        self.dir_list_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
        dir_btns = ctk.CTkFrame(dir_section, fg_color="transparent")
        dir_btns.grid(row=1, column=1, padx=10, pady=10, sticky="n")
        self.add_dir_button = ctk.CTkButton(dir_btns, text="Add Directory...", command=self.add_directory_to_queue, width=120)
        self.add_dir_button.pack(pady=(0, 10), fill="x")
        # Exclude Directories button
        self.exclude_dirs_button = ctk.CTkButton(dir_btns, text="Exclude Directories", command=self.edit_exclude_dirs, width=120)
        self.exclude_dirs_button.pack(pady=(0, 10), fill="x")
        self.clear_queue_button = ctk.CTkButton(dir_btns, text="Clear All", command=self.clear_directory_queue, fg_color="orange red", width=120)
        self.clear_queue_button.pack(pady=(0, 10), fill="x")

        # --- Operation Mode Section ---
        mode_section = ctk.CTkFrame(self.main_frame, fg_color=("#23272e", "#23272e"))
        mode_section.grid(row=1, column=0, sticky="ew", pady=(0, 20))
        mode_section.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(mode_section, text="Operation Mode", font=("Arial", 18, "bold")).grid(row=0, column=0, sticky="w", padx=10, pady=(10, 0))
        self.operation_mode_var = ctk.StringVar(value="Diagnostic Mode (Scan & Report Only)")
        self.mode_dropdown = ctk.CTkComboBox(mode_section, variable=self.operation_mode_var, values=["Diagnostic Mode (Scan & Report Only)", "Production Mode (Scan, Install, & Cleanup Wizard)"], width=400)
        self.mode_dropdown.grid(row=1, column=0, pady=(10, 5), padx=10, sticky="ew")
        self.mode_desc_label = ctk.CTkLabel(mode_section, text="", wraplength=600, justify="left")
        self.mode_desc_label.grid(row=2, column=0, pady=(0, 10), padx=10, sticky="ew")

        # --- Main Action Button ---
        self.start_button = ctk.CTkButton(self.main_frame, text="Run Diagnostic Scan on Queued Directories", command=self.scan_and_install, height=50, font=("Arial", 16))
        self.start_button.grid(row=2, column=0, pady=(0, 25), padx=10, sticky="ew")

        # --- Always Include/Uninstall Section ---
        always_section = ctk.CTkFrame(self.main_frame, fg_color=("#23272e", "#23272e"))
        always_section.grid(row=3, column=0, sticky="ew", pady=(0, 0))
        always_section.grid_columnconfigure(0, weight=1)
        always_section.grid_columnconfigure(1, weight=1)
        # Left: Edit Always Include List
        self.include_button = ctk.CTkButton(always_section, text="Edit Always Include List", command=self.edit_always_include, height=32, width=200, font=("Arial", 12))
        self.include_button.grid(row=0, column=0, padx=(10, 5), pady=10, sticky="w")
        # Right: Edit Always Uninstall List
        self.uninstall_button = ctk.CTkButton(always_section, text="Edit Always Uninstall List", command=self.edit_always_uninstall, height=32, width=200, font=("Arial", 12))
        self.uninstall_button.grid(row=0, column=1, padx=(5, 10), pady=10, sticky="e")
        # Lower right: Close button
        self.close_button = ctk.CTkButton(always_section, text="Close", command=self.destroy, height=32, width=120, font=("Arial", 12), fg_color="#e74c3c", hover_color="#c0392b")
        self.close_button.grid(row=1, column=1, padx=(5, 10), pady=(10, 10), sticky="e")

        # --- Progress Frame (unchanged) ---
        self.progress_frame = ctk.CTkFrame(self)
        self.progress_header = ctk.CTkLabel(self.progress_frame, text="Operation in Progress...", font=("Arial", 20, "bold"))
        self.progress_header.pack(pady=(20, 10))
        self.status_drive_label = ctk.CTkLabel(self.progress_frame, text="Preparing to scan...", font=("Arial", 14))
        self.status_drive_label.pack(pady=5)
        self.status_file_label = ctk.CTkLabel(self.progress_frame, text="", font=("Arial", 12))
        self.status_file_label.pack(pady=5, padx=10)
        self.progress_bar = ctk.CTkProgressBar(self.progress_frame, orientation="horizontal", mode="indeterminate", width=400)
        self.progress_bar.pack(pady=(20, 20))
        self.cancel_button = ctk.CTkButton(self.progress_frame, text="Cancel Operation", command=self.cancel_scan, fg_color="#e74c3c", hover_color="#c0392b")
        self.cancel_button.pack(pady=20)

        self.update_mode_description(self.operation_mode_var.get())
        self.operation_mode_var.trace_add("write", lambda *args: self.update_mode_description(self.operation_mode_var.get()))

    def update_mode_description(self, selected_mode):
        if "Diagnostic" in selected_mode:
            self.mode_desc_label.configure(text="Diagnostic Mode: Scans selected directories and generates a detailed report. NO installations or uninstallations will be performed.")
            self.start_button.configure(text="Run Diagnostic Scan on Queued Directories")
        elif "Production" in selected_mode:
            self.mode_desc_label.configure(text="Production Mode: Scans selected directories, installs missing packages, generates logs, and then launches the Cleanup Wizard.")
            self.start_button.configure(text="Process Queued Directories (Live Mode)")

    def add_directory_to_queue(self):
        directory = filedialog.askdirectory(title="Select Directory Containing Python Scripts")
        if directory:
            # Normalize the path for comparison
            directory = os.path.normpath(directory)
            
            # Simple duplicate check - only prevent exact duplicates
            if directory not in self.queued_directories:
                self.queued_directories.append(directory)
                self.update_queued_dir_display()
            else:
                messagebox.showinfo("Duplicate", "That directory is already in the queue.", parent=self)

    def remove_directory(self, index):
        """Remove a directory at the specified index."""
        if 0 <= index < len(self.queued_directories):
            self.queued_directories.pop(index)
            self.update_queued_dir_display()

    def clear_directory_queue(self):
        if self.queued_directories:
            if messagebox.askyesno("Confirm Clear", "Are you sure you want to clear all directories from the queue?", parent=self):
                self.queued_directories.clear()
                self.update_queued_dir_display()
    
    def update_queued_dir_display(self):
        """Update the directory list display with clickable entries."""
        # Clear existing widgets
        for widget in self.dir_list_frame.winfo_children():
            widget.destroy()

        if not self.queued_directories:
            label = ctk.CTkLabel(
                self.dir_list_frame,
                text="No directories queued. Click 'Add Directory...'",
                text_color="gray"
            )
            label.pack(pady=10, padx=10)
            return

        for i, dir_path in enumerate(self.queued_directories):
            dir_frame = ctk.CTkFrame(self.dir_list_frame)
            dir_frame.pack(fill="x", pady=2, padx=5)
            
            # Create a clickable label for the directory
            dir_label = ctk.CTkLabel(
                dir_frame,
                text=f"{i+1}. {dir_path}",
                anchor="w"
            )
            dir_label.pack(side="left", padx=10, pady=5, fill="x", expand=True)
            
            # Add remove button
            remove_btn = ctk.CTkButton(
                dir_frame,
                text="×",
                width=30,
                command=lambda idx=i: self.remove_directory(idx)
            )
            remove_btn.pack(side="right", padx=5, pady=5)

    def user_is_stdlib_module(self, module_name_to_check):
        module_base_name = module_name_to_check.split('.')[0].lower()
        if module_base_name in ["pillow", "pyside6", "customtkinter", "cv2"]: return False
        if module_base_name in USER_STDLIB_MODULE_LIST: return True
        try:
            spec = importlib.util.find_spec(module_base_name)
            if spec is None: return False
            if spec.origin and "site-packages" in spec.origin.lower().replace("\\", "/"): return False
            if spec.origin and any(lib_path in spec.origin.lower().replace("\\", "/") for lib_path in [os.path.join(sys.prefix, 'lib').lower(), os.path.join(sys.base_prefix, 'lib').lower()]):
                 if "site-packages" not in spec.origin.lower().replace("\\", "/"): return True
        except Exception: pass
        return False

    def user_map_module_to_package(self, module_name_to_map):
        # Handle cases like "from package.module import something" -> map "package.module"
        # For simplicity, we'll map the base first, then if no map, the full.
        # Or, more effectively, map known full dotted paths if they are common.
        # For now, using the original simpler logic for direct import names:
        base_module_name = module_name_to_map.split('.')[0] # e.g. PIL from PIL.Image
        if base_module_name in USER_MODULE_TO_PACKAGE_MAPPING:
            return USER_MODULE_TO_PACKAGE_MAPPING[base_module_name]
        if module_name_to_map in USER_MODULE_TO_PACKAGE_MAPPING: # For direct map like 'cv2'
             return USER_MODULE_TO_PACKAGE_MAPPING[module_name_to_map]
        return base_module_name # Default to base module name if no specific mapping

    def map_and_normalize_imports(self, raw_imports):
        """Map and normalize all import names to PyPI package names, filter out stdlib and junk."""
        current_sys_stdlib = set(sys.stdlib_module_names)
        if sys.version_info >= (3, 10):
            current_sys_stdlib.update(sys.builtin_module_names)
        mapped_pkgs = defaultdict(set)
        for raw_import_name, source_files in raw_imports.items():
            base_module_for_check = raw_import_name.split('.')[0]
            
            # --- HANDLING FOR WHISPER VARIANTS ---
            # Skip all variants of openai whisper to avoid duplicate installations
            if (base_module_for_check.lower().replace("_", "") == "openaiwhisper" or
                base_module_for_check.lower() == "whisper" or 
                base_module_for_check.lower() == "openai-whisper"):
                print(f"[DEBUG] Whisper variant: Mapping '{raw_import_name}' to 'openai-whisper'")
                mapped_pkgs["openai-whisper"].update(source_files)
                continue
                
            print(f"[DEBUG] Checking import: {base_module_for_check.lower()}")
            if base_module_for_check.lower() in JUNK_IMPORT_WORDS:
                print(f"[DEBUG] Ignoring: {base_module_for_check.lower()} (in JUNK_IMPORT_WORDS)")
                continue
            if self.user_is_stdlib_module(base_module_for_check) or base_module_for_check.lower() in current_sys_stdlib:
                continue
            pypi_pkg_name = self.user_map_module_to_package(raw_import_name)
            pypi_pkg_base_name = pypi_pkg_name.split('.')[0].lower()
            if self.user_is_stdlib_module(pypi_pkg_base_name) or pypi_pkg_base_name in current_sys_stdlib:
                continue
            if pypi_pkg_name.startswith('_') and pypi_pkg_name not in {"_cffi_backend"}: continue
            if not re.match(r"^[a-zA-Z0-9_.-]+$", pypi_pkg_name): continue
            mapped_pkgs[pypi_pkg_name.lower()].update(source_files)
        # Always include packages (map and normalize to PyPI names)
        for pkg in self.always_include:
            mapped_name = self.user_map_module_to_package(pkg).lower()
            if mapped_name not in mapped_pkgs:
                mapped_pkgs[mapped_name] = {'[AlwaysInclude]'}
        return mapped_pkgs

    def scan_and_install(self):
        if not self.queued_directories:
            messagebox.showerror("Error", "No directories selected for scanning. Please add directories to the queue.", parent=self)
            self.restart_app() # Ensure UI is reset
            return
        self.main_frame.pack_forget()
        self.progress_frame.pack(fill="both", expand=True, padx=20, pady=20)
        self.progress_bar.start()
        self.cancel_event.clear()  # Reset cancel event
        raw_imports_to_files_map = defaultdict(set)
        installed_now_for_report = []
        is_diagnostic_run = "Diagnostic" in self.operation_mode_var.get()
        final_status_message = "Operation completed."
        try:
            self.update_status("Initial scan: Discovering Python files...")
            all_py_files = set()
            scanned_dirs = set()
            file_counts = defaultdict(int)
            duplicate_dirs = set()
            # Expand environment variables in exclude_dirs
            expanded_exclude_dirs = [os.path.expandvars(path) for path in self.exclude_dirs]
            expanded_exclude_dirs = [os.path.normpath(path) for path in expanded_exclude_dirs]
            for dir_idx, directory_to_scan in enumerate(self.queued_directories):
                if self.cancel_event.is_set():
                    self.on_complete(defaultdict(set), [], [], "Operation cancelled.", is_diagnostic_run)
                    return
                directory_to_scan = os.path.normpath(directory_to_scan)
                self.update_status(f"Scanning Directory {dir_idx+1}/{len(self.queued_directories)}", f"...{directory_to_scan[-50:]}")
                for root, dirs, files_in_dir in os.walk(directory_to_scan, topdown=True):
                    if self.cancel_event.is_set():
                        self.on_complete(defaultdict(set), [], [], "Operation cancelled.", is_diagnostic_run)
                        return
                    # Exclude subfolders by name (as before)
                    dirs[:] = [d for d in dirs if d.lower() not in {name.lower() for name in EXCLUDED_SUBFOLDER_NAMES}]
                    # Exclude by full path (new logic)
                    root_norm = os.path.normpath(root)
                    if any(root_norm.startswith(excl) for excl in expanded_exclude_dirs):
                        continue
                    if root in scanned_dirs:
                        duplicate_dirs.add(root)
                        continue
                    scanned_dirs.add(root)
                    for file_name in files_in_dir:
                        if file_name.lower().endswith(('.py', '.pyw')):
                            full_path = os.path.join(root, file_name)
                            all_py_files.add(full_path)
                            file_counts[directory_to_scan] += 1
            if not all_py_files:
                self.on_complete(defaultdict(set), [], [], "No Python files found in selected directories.", is_diagnostic_run)
                return
            summary_lines = ["File Discovery Summary:"]
            total_files = len(all_py_files)
            total_dirs = len(scanned_dirs)
            skipped_dirs = len(duplicate_dirs)
            for dir_path, count in file_counts.items():
                summary_lines.append(f"\n{dir_path}:")
                summary_lines.append(f"  - Found {count} Python file(s)")
            summary_lines.append(f"\nTotal unique Python files found: {total_files}")
            summary_lines.append(f"Total directories scanned: {total_dirs}")
            if skipped_dirs > 0:
                summary_lines.append(f"Directories skipped (already scanned): {skipped_dirs}")
            summary_text = "\n".join(summary_lines)
            if not messagebox.askyesno(
                "File Discovery Complete",
                f"{summary_text}\n\nProceed with analysis?",
                parent=self
            ):
                self.on_complete(defaultdict(set), [], [], "Operation cancelled by user.", is_diagnostic_run)
                return
            self.update_status(f"Analyzing imports from {total_files} files...")
            for i, file_path in enumerate(all_py_files):
                if self.cancel_event.is_set():
                    self.on_complete(defaultdict(set), [], [], "Operation cancelled.", is_diagnostic_run)
                    return
                self.update_status(
                    f"Analyzing file {i+1}/{total_files} ({(i+1)/total_files*100:.1f}%)",
                    f"...{file_path}"
                )
                try:
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f_content:
                        content = f_content.read()
                        for pattern in USER_IMPORT_PATTERNS:
                            matches = re.finditer(pattern, content, re.MULTILINE)
                            for match in matches:
                                module_name_match = match.group(1)
                                if module_name_match:
                                    raw_imports_to_files_map[module_name_match].add(file_path)
                except (PermissionError, FileNotFoundError, OSError) as e:
                    print(f"Error reading {file_path}: {e}")
                    continue
            # --- Consistent mapping/normalization for all steps ---
            valid_pypi_pkgs_to_files_map = self.map_and_normalize_imports(raw_imports_to_files_map)
            needed_pypi_pkgs_set = set(valid_pypi_pkgs_to_files_map.keys())
            current_installed_system_pkgs_set = {pkg.lower() for pkg in self.get_current_installed_pypi_packages()}
            missing_packages_to_install = sorted(list(needed_pypi_pkgs_set - current_installed_system_pkgs_set))
            if is_diagnostic_run:
                self.generate_diagnostic_report(raw_imports_to_files_map, valid_pypi_pkgs_to_files_map, current_installed_system_pkgs_set, missing_packages_to_install)
                return
            if not missing_packages_to_install:
                final_status_message = "System is up to date. No new packages to install based on selected directories."
                self.on_complete(valid_pypi_pkgs_to_files_map, [], list(current_installed_system_pkgs_set), final_status_message, is_diagnostic_run)
                return
            total_missing = len(missing_packages_to_install)
            for i, pkg_to_install in enumerate(missing_packages_to_install):
                if self.cancel_event.is_set():
                    self.on_complete(valid_pypi_pkgs_to_files_map, installed_now_for_report, list(current_installed_system_pkgs_set), "Operation cancelled.", is_diagnostic_run)
                    return
                self.update_status(f"Installing package {i+1}/{total_missing}", f"pip install {pkg_to_install}")
                print(f"[DEBUG] Attempting to install: {pkg_to_install}")  # DEBUG LINE
                try:
                    startupinfo = None
                    if platform.system() == "Windows":
                        startupinfo = subprocess.STARTUPINFO()
                        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    subprocess.check_call([sys.executable, '-m', 'pip', 'install', pkg_to_install], stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT, startupinfo=startupinfo)
                    installed_now_for_report.append(pkg_to_install)
                except subprocess.CalledProcessError as e:
                    installed_now_for_report.append(f"{pkg_to_install} (FAILED: pip error {e.returncode})")
            num_successful_installs = sum(1 for p in installed_now_for_report if "(FAILED" not in p)
            final_status_message = f"Successfully installed {num_successful_installs} of {len(installed_now_for_report)} attempted package(s)."
        except Exception as e:
            final_status_message = f"An unexpected error occurred: {str(e)}"
        final_installed_system_pkgs_set = {pkg.lower() for pkg in self.get_current_installed_pypi_packages()}
        self.on_complete(valid_pypi_pkgs_to_files_map, installed_now_for_report, list(final_installed_system_pkgs_set), final_status_message, is_diagnostic_run)

    def generate_diagnostic_report(self, raw_imports, valid_pkgs, installed_set, missing_list):
        self.progress_bar.stop()
        downloads_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
        if not os.path.exists(downloads_dir):
            try: os.makedirs(downloads_dir)
            except OSError: downloads_dir = os.path.expanduser('~')
        report_path = os.path.join(downloads_dir, "diagnostic_report.log")
        with open(report_path, 'w', encoding='utf-8') as f_report: 
            f_report.write("--- SCRIPT LOGIC DIAGNOSTIC REPORT ---\n")
            f_report.write(f"Report generated on: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f_report.write(f"Scan limited to user-selected directories.\n")
            # 1. Paths current script and package roster
            f_report.write("\n\n--- SECTION 1: Paths current script and package roster ---\n")
            f_report.write("Format: import_name\n----------------------------------------\n[path to each file using this import]\n")
            if raw_imports:
                for imp, files in sorted(raw_imports.items()):
                    f_report.write(f"{imp}\n")
                    f_report.write("-"*40 + "\n")
                    for file_path in sorted(files):
                        f_report.write(f"{file_path}\n")
                    f_report.write("\n")
            else:
                f_report.write("No raw imports were collected.\n")
            # 2. Packages installed with script dependencies
            f_report.write("\n\n--- SECTION 2: Packages installed with script dependencies ---\n")
            f_report.write("Format: package_name\n----------------------------------------\n[path to each file using this package]\n")
            # Only include packages that are both installed and used by scripts
            installed_and_used = set(valid_pkgs.keys()) & set(installed_set)
            if installed_and_used:
                for pkg in sorted(installed_and_used):
                    files = valid_pkgs.get(pkg, set())
                    f_report.write(f"{pkg}\n")
                    f_report.write("-"*40 + "\n")
                    for file_path in sorted(files):
                        f_report.write(f"{file_path}\n")
                    f_report.write("\n")
            else:
                f_report.write("No installed packages are used by your scripts.\n")
            # 3. Packages Needing Deletion (They are not needed)
            f_report.write("\n\n--- SECTION 3: Packages Needing Deletion (They are not needed) ---\n")
            f_report.write("Format: package_name\n")
            # Exclude protected packages, always_include packages, and packages in always_uninstall
            staged_for_deletion = set(installed_set) - set(valid_pkgs.keys()) - PROTECTED_PACKAGES - self.always_include
            if staged_for_deletion:
                f_report.write("\n".join(sorted(list(staged_for_deletion))))
            else:
                f_report.write("No packages are staged for deletion.\n")
            # 4. Packages Needing Installation (They are missing)
            f_report.write("\n\n--- SECTION 4: Packages Needing Installation (They are missing) ---\n")
            f_report.write("Format: package_name\n")
            # Filter out packages that are in always_uninstall
            missing_list = [pkg for pkg in missing_list if pkg.lower() not in {p.lower() for p in self.always_uninstall}]
            if missing_list:
                f_report.write("\n".join(sorted(missing_list)))
            else:
                f_report.write("No packages were identified as missing.\n")
        final_message = f"Diagnostic Mode complete.\nA detailed report has been saved to:\n{report_path}"
        # Auto-open the Downloads directory after report is saved
        try:
            if platform.system() == "Windows":
                os.startfile(downloads_dir)
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", downloads_dir])
            else:
                subprocess.Popen(["xdg-open", downloads_dir])
        except Exception as e:
            print(f"Could not open directory: {e}")
        messagebox.showinfo("Diagnostic Complete", final_message, parent=self) 
        self.restart_app()

    def on_complete(self, valid_pkgs_with_sources_map, newly_installed_list, final_system_pkgs_list, final_message_status, is_diagnostic_run):
        if is_diagnostic_run: 
            self.restart_app() 
            return

        self.progress_bar.stop()
        downloads_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
        if not os.path.exists(downloads_dir):
            try: os.makedirs(downloads_dir)
            except OSError: downloads_dir = os.path.expanduser('~')

        # --- Install Results Popup ---
        already_present = []
        newly_installed = []
        failed = []
        for item in newly_installed_list:
            if "(FAILED" in item:
                failed.append(item)
            else:
                newly_installed.append(item)
        # Figure out which needed packages were already present
        needed_pkgs = set(valid_pkgs_with_sources_map.keys())
        for pkg in needed_pkgs:
            if pkg not in newly_installed and pkg in final_system_pkgs_list:
                already_present.append(pkg)
        summary_lines = []
        if newly_installed:
            summary_lines.append("Newly Installed Packages:\n" + "\n".join(newly_installed) + "\n")
        if already_present:
            summary_lines.append("Already Present Packages:\n" + "\n".join(already_present) + "\n")
        if failed:
            summary_lines.append("Failed Installs:\n" + "\n".join(failed) + "\n")
        summary_text = "\n".join(summary_lines) if summary_lines else "No packages were newly installed or attempted."

        def export_to_downloads():
            export_path = os.path.join(downloads_dir, "install_summary.log")
            with open(export_path, 'w', encoding='utf-8') as f:
                f.write(summary_text)
            messagebox.showinfo("Export Complete", f"Summary exported to:\n{export_path}", parent=dialog)

        dialog = ctk.CTkToplevel(self)
        dialog.title("Install Results Summary")
        dialog.geometry("600x400"); dialog.transient(self); dialog.grab_set()
        ctk.CTkLabel(dialog, text="Install Results", font=("Arial", 16, "bold")).pack(pady=(10, 5))
        textbox = ctk.CTkTextbox(dialog, wrap="word", height=250, width=550)
        textbox.pack(pady=10, padx=20, fill="both", expand=True)
        textbox.insert("1.0", summary_text)
        textbox.configure(state="normal")  # Allow copy
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(pady=10)
        export_btn = ctk.CTkButton(btn_frame, text="Export to Downloads", command=export_to_downloads)
        export_btn.pack(side="left", padx=10)
        
        # Modified to launch cleanup wizard after closing the dialog
        def continue_to_cleanup():
            dialog.destroy()
            self.launch_cleanup_wizard(set(valid_pkgs_with_sources_map.keys()), final_system_pkgs_list)
            
        def cancel_operation():
            dialog.destroy()
            self.restart_app()
            
        if is_diagnostic_run:
            close_btn = ctk.CTkButton(btn_frame, text="Close", command=cancel_operation)
        else:
            close_btn = ctk.CTkButton(btn_frame, text="Continue to Cleanup...", command=continue_to_cleanup)
            cancel_btn = ctk.CTkButton(btn_frame, text="Cancel", command=cancel_operation, fg_color="#e74c3c", hover_color="#c0392b")
            cancel_btn.pack(side="left", padx=10)
            
        close_btn.pack(side="left", padx=10)

        # --- Existing log writing for requirements and installs ---
        needed_log_path = os.path.join(downloads_dir, "systemlevel_package_requirements.log")
        with open(needed_log_path, 'w', encoding='utf-8') as f_needed:
            f_needed.write("Needed Packages (with source scripts)\n" + "=" * 38 + "\n")
            if valid_pkgs_with_sources_map:
                for pkg, sources in sorted(valid_pkgs_with_sources_map.items()):
                    f_needed.write(f"{pkg} (found in: {', '.join(sorted(list(sources)))})\n")
            else: f_needed.write("No third-party packages identified as needed.\n")
        installed_log_path = os.path.join(downloads_dir, "systemlevel_package_newlyinstalled.log")
        with open(installed_log_path, 'w', encoding='utf-8') as f_installed:
            f_installed.write("Installed Packages (with source scripts)\n" + "=" * 40 + "\n")
            if newly_installed_list:
                for item in newly_installed_list:
                    pkg_name_for_lookup = item.split(" (FAILED")[0]
                    sources = valid_pkgs_with_sources_map.get(pkg_name_for_lookup, set())
                    f_installed.write(f"{item} (source scripts: {', '.join(sorted(list(sources))) if sources else 'N/A'})\n")
            else: f_installed.write("No packages were newly installed or attempted.\n")

    def launch_cleanup_wizard(self, needed_pypi_pkgs_set, all_system_pkgs_list):
        # Use the same normalization and set logic for deletion
        all_installed_with_deps_info = self.get_installed_packages_with_deps()
        all_system_pkgs_set = {pkg.lower() for pkg in all_system_pkgs_list}
        needed_pypi_pkgs_set = {pkg.lower() for pkg in needed_pypi_pkgs_set}
        # Exclude protected packages, always_include packages, and include always_uninstall packages
        potentially_unused = all_system_pkgs_set - needed_pypi_pkgs_set - PROTECTED_PACKAGES - self.always_include
        # Add packages from always_uninstall list
        always_uninstall_lower = {pkg.lower() for pkg in self.always_uninstall}
        potentially_unused.update(always_uninstall_lower)
        if not potentially_unused:
            messagebox.showinfo("Cleanup Not Needed", "No unused packages (excluding protected) found to clean up.", parent=self)
            self.restart_app()
            return
        display_text_lines = []
        required_by_map = defaultdict(set)
        for pkg, deps in all_installed_with_deps_info.items():
            for dep in deps: required_by_map[dep].add(pkg)
        unused_dependency_tree = defaultdict(set)
        for pkg in sorted(list(potentially_unused)):
            deps_of_pkg = all_installed_with_deps_info.get(pkg, set())
            for dep in deps_of_pkg:
                if dep in potentially_unused and not (required_by_map[dep] - potentially_unused - needed_pypi_pkgs_set):
                     unused_dependency_tree[pkg].add(dep)
        processed_for_display = set()
        sorted_top_level_unused = sorted(list(potentially_unused - set(d for deps_list in unused_dependency_tree.values() for d in deps_list)))
        for pkg in sorted_top_level_unused:
            if pkg in processed_for_display: continue
            # Mark packages from always_uninstall with a special indicator
            if pkg in always_uninstall_lower:
                display_text_lines.append(f"{pkg} [Always Uninstall]")
            else:
                display_text_lines.append(f"{pkg}")
            processed_for_display.add(pkg)
            if pkg in unused_dependency_tree:
                for dep in sorted(list(unused_dependency_tree[pkg])):
                    if dep in processed_for_display: continue
                    if dep in always_uninstall_lower:
                        display_text_lines.append(f"  - {dep} [Always Uninstall]")
                    else:
                        display_text_lines.append(f"  - {dep}")
                    processed_for_display.add(dep)
        for pkg in sorted(list(potentially_unused)):
            if pkg not in processed_for_display:
                if pkg in always_uninstall_lower:
                    display_text_lines.append(f"{pkg} [Always Uninstall]")
                else:
                    display_text_lines.append(f"{pkg}")
        display_text = "\n".join(display_text_lines)
        dialog = ctk.CTkToplevel(self)
        dialog.title("Cleanup Wizard (Stage 1): Review & Edit")
        dialog.geometry("600x600"); dialog.transient(self); dialog.grab_set()
        ctk.CTkLabel(dialog, text="Potentially Unused Packages", font=("Arial", 16, "bold")).pack(pady=(10, 5))
        ctk.CTkLabel(dialog, text="Review the list. Dependencies are indented. Delete lines for packages you want to KEEP.\nPackages marked [Always Uninstall] are from your Always Uninstall list.", wraplength=550).pack(pady=5)
        textbox = ctk.CTkTextbox(dialog, wrap="none", height=400, width=550)
        textbox.pack(pady=10, padx=20, fill="both", expand=True)
        textbox.insert("1.0", display_text)
        button_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        button_frame.pack(pady=10)
        def proceed_stage1():
            text_content = textbox.get("1.0", "end-1c")
            packages_to_consider_uninstall = sorted(list(set(line.strip().replace("- ", "").replace(" [Always Uninstall]", "") for line in text_content.splitlines() if line.strip() and not line.strip().startswith("#"))))
            dialog.destroy()
            if packages_to_consider_uninstall:
                self.launch_uninstall_confirmation(packages_to_consider_uninstall)
            else:
                messagebox.showinfo("No Selection", "No packages remained after editing.", parent=self)
                self.restart_app()
        ctk.CTkButton(button_frame, text="Continue to Uninstall...", command=proceed_stage1).pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="Leave As Is (Cancel)", command=lambda: [dialog.destroy(), self.restart_app()], fg_color="#e74c3c", hover_color="#c0392b").pack(side="left", padx=10)

    def launch_uninstall_confirmation(self, packages_to_uninstall_list):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Cleanup Wizard (Stage 2): Final Confirmation")
        dialog.geometry("600x600"); dialog.transient(self); dialog.grab_set()
        ctk.CTkLabel(dialog, text="Final Confirmation: Uninstall Packages", font=("Arial", 16, "bold")).pack(pady=(10,5))
        ctk.CTkLabel(dialog, text="Check boxes for packages to permanently uninstall. Uncheck to keep.", wraplength=550).pack(pady=5)
        scroll_frame = ctk.CTkScrollableFrame(dialog, label_text="Packages for Uninstallation")
        scroll_frame.pack(pady=10, padx=20, fill="both", expand=True)
        checkbox_vars = {}
        checkbox_widgets = []
        for pkg in packages_to_uninstall_list:
            var = ctk.StringVar(value="on") 
            cb = ctk.CTkCheckBox(scroll_frame, text=pkg, variable=var, onvalue="on", offvalue="off")
            cb.pack(anchor="w", padx=10, pady=2)
            checkbox_vars[pkg] = var
            checkbox_widgets.append(cb)
        status_label = ctk.CTkLabel(dialog, text="")
        status_label.pack(pady=5)
        def do_final_uninstall():
            # Disable all controls during uninstall
            try:
                for cb in checkbox_widgets:
                    if cb.winfo_exists():  # Check if widget still exists
                        cb.configure(state="disabled")
                uninstall_button.configure(state="disabled")
                cancel_button.configure(state="disabled")
            except Exception as e:
                print(f"Widget error during disable: {e}")
                
            final_packages_to_remove = [pkg for pkg, var in checkbox_vars.items() if var.get() == "on"]
            if not final_packages_to_remove:
                messagebox.showinfo("No Selection", "No packages were selected for uninstallation.", parent=dialog)
                return
            def uninstall_thread_target():
                uninstalled_log, failed_log = [], []
                total_to_remove = len(final_packages_to_remove)
                for i, pkg_name in enumerate(final_packages_to_remove):
                    if self.cancel_event.is_set(): 
                        try:
                            if status_label.winfo_exists():
                                status_label.configure(text="Uninstallation cancelled.")
                        except Exception:
                            pass
                        break
                    try:
                        if status_label.winfo_exists():
                            status_label.configure(text=f"Uninstalling {i+1}/{total_to_remove}: {pkg_name}...")
                    except Exception:
                        pass
                    
                    try:
                        startupinfo = None
                        if platform.system() == "Windows":
                            startupinfo = subprocess.STARTUPINFO()
                            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                        subprocess.check_call([sys.executable, "-m", "pip", "uninstall", "-y", pkg_name], stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT, startupinfo=startupinfo)
                        uninstalled_log.append(pkg_name)
                    except subprocess.CalledProcessError: 
                        failed_log.append(pkg_name)
                    except Exception as e:
                        print(f"Error uninstalling {pkg_name}: {e}")
                        failed_log.append(f"{pkg_name} (error)")
                        
                report_message = f"Uninstallation process finished.\nSuccessfully uninstalled: {len(uninstalled_log)}\n"
                if uninstalled_log: report_message += f" ({', '.join(uninstalled_log)})\n"
                report_message += f"Failed to uninstall: {len(failed_log)}\n"
                if failed_log: report_message += f" ({', '.join(failed_log)})\n"
                
                try:
                    if status_label.winfo_exists():
                        status_label.configure(text="Uninstallation complete!")
                    messagebox.showinfo("Uninstallation Complete", report_message, parent=self)
                except Exception:
                    # Fallback if dialog was closed
                    messagebox.showinfo("Uninstallation Complete", report_message)
                
                try:
                    if dialog.winfo_exists():
                        dialog.destroy()
                except Exception:
                    pass
                    
                self.restart_app()
            threading.Thread(target=uninstall_thread_target, daemon=True).start()
        button_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        button_frame.pack(pady=10)
        uninstall_button = ctk.CTkButton(button_frame, text="Uninstall Selected Packages", command=do_final_uninstall, fg_color="red")
        uninstall_button.pack(side="left", padx=10)
        cancel_button = ctk.CTkButton(button_frame, text="Cancel Cleanup", command=lambda: [dialog.destroy(), self.restart_app()])
        cancel_button.pack(side="left", padx=10)

    def cancel_scan(self):
        """Handle cancellation of ongoing operations."""
        self.cancel_event.set()
        self.status_drive_label.configure(text="Cancelling operation...")
        self.status_file_label.configure(text="Please wait...")
        self.cancel_button.configure(state="disabled")
        
    def restart_app(self):
        self.progress_frame.pack_forget()
        self.main_frame.pack(fill="both", expand=True, padx=30, pady=30)
        self.start_button.configure(state="normal")
        self.mode_dropdown.configure(state="normal") 
        self.status_drive_label.configure(text="Preparing to scan...") # Reset status for next run
        self.status_file_label.configure(text="")

    def get_installed_packages_with_deps(self):
        """Get a mapping of installed packages to their dependencies."""
        try:
            startupinfo = None
            if platform.system() == "Windows":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            
            # Get pip list output
            result = subprocess.check_output(
                [sys.executable, "-m", "pip", "list", "--format=json"],
                stderr=subprocess.STDOUT,
                startupinfo=startupinfo
            ).decode('utf-8')
            
            # Parse the JSON output
            installed_packages = json.loads(result)
            package_to_deps = {}
            
            # For each installed package, get its dependencies
            for pkg in installed_packages:
                pkg_name = pkg['name']
                try:
                    deps_result = subprocess.check_output(
                        [sys.executable, "-m", "pip", "show", pkg_name],
                        stderr=subprocess.STDOUT,
                        startupinfo=startupinfo
                    ).decode('utf-8')
                    
                    # Parse dependencies from pip show output
                    deps = set()
                    for line in deps_result.split('\n'):
                        if line.startswith('Requires:'):
                            deps.update(d.strip() for d in line[9:].split(',') if d.strip())
                    
                    package_to_deps[pkg_name] = deps
                except subprocess.CalledProcessError:
                    package_to_deps[pkg_name] = set()
            
            return package_to_deps
        except Exception as e:
            print(f"Error getting package dependencies: {e}")
            return {}

    def get_current_installed_pypi_packages(self):
        """Get a set of currently installed PyPI packages."""
        try:
            startupinfo = None
            if platform.system() == "Windows":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            
            # Get pip list output
            result = subprocess.check_output(
                [sys.executable, "-m", "pip", "list", "--format=json"],
                stderr=subprocess.STDOUT,
                startupinfo=startupinfo
            ).decode('utf-8')
            
            # Parse the JSON output and extract package names
            installed_packages = json.loads(result)
            return {pkg['name'].lower() for pkg in installed_packages}
        except Exception as e:
            print(f"Error getting installed packages: {e}")
            return set()

    def update_status(self, drive_text, file_text=""):
        """Update the status labels with current operation information."""
        self.status_drive_label.configure(text=drive_text)
        self.status_file_label.configure(text=file_text)
        self.update_idletasks()  # Force GUI update

    def load_settings_ini(self):
        config = configparser.RawConfigParser()
        if os.path.exists(SETTINGS_INI):
            config.read(SETTINGS_INI)
            self.always_include = set(config.get('AlwaysInclude', 'packages', fallback='').splitlines())
            self.always_uninstall = set(config.get('AlwaysUninstall', 'packages', fallback='').splitlines())
            self.exclude_dirs = set(config.get('ExcludeDirs', 'paths', fallback='').splitlines())
        else:
            # Pre-populate with defaults if ini is missing
            self.always_include = {'PySide6', 'OpenAIWhisper'}
            self.always_uninstall = set()
            self.exclude_dirs = set(DEFAULT_EXCLUDE_DIRS)
            self.save_settings_ini()
        # If the ini exists but AlwaysInclude is empty, pre-populate
        if not self.always_include:
            self.always_include = {'PySide6', 'OpenAIWhisper'}
            self.save_settings_ini()
        if not self.exclude_dirs:
            self.exclude_dirs = set(DEFAULT_EXCLUDE_DIRS)
            self.save_settings_ini()

    def save_settings_ini(self):
        config = configparser.RawConfigParser()
        config['AlwaysInclude'] = {'packages': '\n'.join(sorted(self.always_include))}
        config['AlwaysUninstall'] = {'packages': '\n'.join(sorted(self.always_uninstall))}
        config['ExcludeDirs'] = {'paths': '\n'.join(sorted(self.exclude_dirs))}
        with open(SETTINGS_INI, 'w', encoding='utf-8') as f:
            config.write(f)

    def edit_always_include(self):
        self._edit_list_dialog('Always Include Packages', self.always_include, self.save_always_include)

    def edit_always_uninstall(self):
        self._edit_list_dialog('Always Uninstall Packages', self.always_uninstall, self.save_always_uninstall)

    def _edit_list_dialog(self, title, pkg_set, save_callback):
        dialog = ctk.CTkToplevel(self)
        dialog.title(title)
        dialog.geometry("400x520"); dialog.transient(self); dialog.grab_set()
        ctk.CTkLabel(dialog, text=title, font=("Arial", 14, "bold")).pack(pady=10)
        textbox = ctk.CTkTextbox(dialog, wrap="none", height=300, width=350)
        textbox.pack(pady=10, padx=20, fill="both", expand=True)
        textbox.insert("1.0", '\n'.join(sorted(pkg_set)))
        def save():
            lines = [line.strip() for line in textbox.get("1.0", "end-1c").splitlines() if line.strip()]
            pkg_set.clear(); pkg_set.update(lines)
            save_callback()
            dialog.destroy()
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(side="bottom", fill="x", pady=(10, 15), padx=20)
        save_btn = ctk.CTkButton(btn_frame, text="Save", command=save, height=36, width=120, font=("Arial", 12))
        save_btn.pack(side="left", expand=True, fill="x", padx=(0, 10))
        cancel_btn = ctk.CTkButton(btn_frame, text="Cancel", command=dialog.destroy, height=36, width=120, font=("Arial", 12), fg_color="#444", hover_color="#222")
        cancel_btn.pack(side="right", expand=True, fill="x", padx=(10, 0))

    def save_always_include(self):
        self.save_settings_ini()
    def save_always_uninstall(self):
        self.save_settings_ini()

    def edit_exclude_dirs(self):
        self._edit_exclude_dirs_dialog('Edit Excluded Directories', self.exclude_dirs, self.save_exclude_dirs)

    def _edit_exclude_dirs_dialog(self, title, dir_set, save_callback):
        dialog = ctk.CTkToplevel(self)
        dialog.title(title)
        dialog.geometry("500x520"); dialog.transient(self); dialog.grab_set()
        ctk.CTkLabel(dialog, text=title, font=("Arial", 14, "bold")).pack(pady=10)
        textbox = ctk.CTkTextbox(dialog, wrap="none", height=350, width=450)
        textbox.pack(pady=10, padx=20, fill="both", expand=True)
        textbox.insert("1.0", '\n'.join(sorted(dir_set)))
        def save():
            lines = [line.strip() for line in textbox.get("1.0", "end-1c").splitlines() if line.strip()]
            dir_set.clear(); dir_set.update(lines)
            save_callback()
            dialog.destroy()
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(side="bottom", fill="x", pady=(10, 15), padx=20)
        save_btn = ctk.CTkButton(btn_frame, text="Save", command=save, height=36, width=120, font=("Arial", 12))
        save_btn.pack(side="left", expand=True, fill="x", padx=(0, 10))
        cancel_btn = ctk.CTkButton(btn_frame, text="Cancel", command=dialog.destroy, height=36, width=120, font=("Arial", 12), fg_color="#444", hover_color="#222")
        cancel_btn.pack(side="right", expand=True, fill="x", padx=(10, 0))

    def save_exclude_dirs(self):
        self.save_settings_ini()

def is_admin():
    try:
        if platform.system() == "Windows":
            return ctypes.windll.shell32.IsUserAnAdmin() != 0
        else: return os.geteuid() == 0
    except Exception: return False

if __name__ == "__main__":
    print("Creating RequirementsDoctor instance...")
    app = RequirementsDoctor()
    print("Starting mainloop...")
    app.mainloop()
    print("Mainloop ended.")
