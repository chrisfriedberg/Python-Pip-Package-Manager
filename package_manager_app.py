import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, Listbox, Frame, simpledialog, Toplevel, Checkbutton, BooleanVar
import subprocess
import sys
import os
import threading
import json
import math
import locale
import queue
import site
import webbrowser  # For opening PyPI search
from PIL import Image, ImageTk  # <-- Add this import
import datetime
from venv_creator import VenvCreatorDialog  # Add this import
# Add pystray import
try:
    import pystray
except ImportError:
    pystray = None
    print("pystray not installed. System tray integration will not work. Run 'pip install pystray pillow'.")

# For normalization warning
try:
    from packaging.utils import canonicalize_name
    PACKAGING_UTILS_AVAILABLE = True
except ImportError:
    PACKAGING_UTILS_AVAILABLE = False
    print("Warning: 'packaging' not found. Run 'pip install packaging' for better name normalization.")
    def canonicalize_name(name):  # Crude fallback
        return name.lower().replace("_", "-").replace(".", "-")

try:
    from CTkToolTip import CTkToolTip
except ImportError:
    class CTkToolTip:
        def __init__(self, widget, message, **kwargs):
            pass

# For package listing (replacing pkg_resources)
try:
    from importlib import metadata as importlib_metadata
except ImportError:
    importlib_metadata = None

# Platform detection
IS_WINDOWS = sys.platform.startswith("win32")
IS_MAC = sys.platform == "darwin"
IS_LINUX = sys.platform.startswith("linux")

# CREATE_NO_WINDOW flag
if IS_WINDOWS:
    from subprocess import CREATE_NO_WINDOW
else:
    CREATE_NO_WINDOW = 0

# Python environment variables
PYTHON_EXECUTABLE = sys.executable
PYTHON_VERSION = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
PYTHON_BASE_DIR = os.path.dirname(PYTHON_EXECUTABLE)
PYTHON_SCRIPTS_DIR = os.path.join(PYTHON_BASE_DIR, "Scripts" if IS_WINDOWS else "bin")
PYTHON_SITE_PACKAGES = site.getsitepackages()[0] if site.getsitepackages() else None

# Platform-specific app data directory
if IS_MAC:
    BASE_DATA_DIR = os.path.expanduser("~/Library/Application Support")
elif IS_WINDOWS:
    BASE_DATA_DIR = os.getenv('APPDATA') or os.path.expanduser("~\\AppData\\Roaming")
else:
    BASE_DATA_DIR = os.path.expanduser("~/.local/share")

APP_NAME = "Python Global Package Manager (V1.19)"
APP_DATA_DIR = os.path.join(BASE_DATA_DIR, "Python_Global_Package_Manager")
LOG_FILE = os.path.join(APP_DATA_DIR, "package_manager.log")

# Ensure app data directory exists
try:
    os.makedirs(APP_DATA_DIR, exist_ok=True)
except Exception as e:
    print(f"Error creating app data directory: {e}")

# --- Constants ---
GEOMETRY = "950x700"
NUM_COLUMNS = 5
LISTBOX_DEFAULT_HEIGHT = 20
TERMINAL_FONT = ("Consolas", 10)
TERMINAL_HEIGHT = 8
DESCRIPTION_BOX_HEIGHT = 170
LISTBOX_FONT = ("Consolas", 11)  # Monospace for alignment

MENU_BG_COLOR = "#3C3C3C"
MENU_FG_COLOR = "#FFFFFF"
MENU_ACTIVE_BG_COLOR = "#505050"
MENU_FONT = ("Arial", 10)

TITLE_BAR_COLOR = "#1A73E8"

PACKAGE_NAME_CACHE_FILE = os.path.join(APP_DATA_DIR, "common_packages.json")
ICON_LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_icon_path.txt")

# --- Helper Functions ---
def get_pip_path():
    """Get the path to pip executable, preferring the global Python installation."""
    # First try to find pip in the global Python installation
    if IS_WINDOWS:
        # On Windows, check common installation paths
        python_paths = [
            os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Programs', 'Python'),
            os.path.join(os.environ.get('PROGRAMFILES', ''), 'Python'),
            os.path.join(os.environ.get('PROGRAMFILES(X86)', ''), 'Python'),
            r'C:\Python',
            r'C:\Python3',
            r'C:\Python39',
            r'C:\Python310',
            r'C:\Python311',
            r'C:\Python312'
        ]
        for base_path in python_paths:
            if os.path.exists(base_path):
                for version_dir in os.listdir(base_path):
                    if version_dir.startswith('Python'):
                        pip_path = os.path.join(base_path, version_dir, 'Scripts', 'pip.exe')
                        if os.path.exists(pip_path) and os.access(pip_path, os.X_OK):
                            return pip_path
    else:
        # On Unix-like systems, check common paths
        pip_paths = [
            '/usr/bin/pip3',
            '/usr/local/bin/pip3',
            '/usr/bin/pip',
            '/usr/local/bin/pip'
        ]
        for path in pip_paths:
            if os.path.exists(path) and os.access(path, os.X_OK):
                return path

    # If global pip not found, try using the current Python's pip
    try:
        if IS_WINDOWS:
            result = subprocess.run(['where', 'pip'], capture_output=True, text=True, check=False)
        else:
            result = subprocess.run(['which', 'pip'], capture_output=True, text=True, check=False)
        
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout.strip().split('\n')[0]
    except Exception:
        pass

    return None

def get_python_env_info():
    """Get information about the current Python environment."""
    return {
        'executable': PYTHON_EXECUTABLE,
        'version': PYTHON_VERSION,
        'base_dir': PYTHON_BASE_DIR,
        'scripts_dir': PYTHON_SCRIPTS_DIR,
        'site_packages': PYTHON_SITE_PACKAGES,
        'platform': sys.platform,
        'is_venv': hasattr(sys, 'real_prefix') or (hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix),
        'venv_path': sys.prefix if hasattr(sys, 'real_prefix') or (hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix) else None
    }

def run_pip_command_live(command, output_queue, timeout=300):
    pip_path = get_pip_path()
    if not pip_path:
        output_queue.put(('stderr', "[PIP_NOT_FOUND] Could not locate pip executable.\n"))
        return "PIP_NOT_FOUND"
    full_command = [pip_path] + command
    try:
        process = subprocess.Popen(
            full_command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
            encoding=locale.getpreferredencoding(False) or 'utf-8', errors='replace',
            bufsize=1, shell=False, env=os.environ.copy(), creationflags=CREATE_NO_WINDOW if IS_WINDOWS else 0
        )
        def reader_thread(pipe, stream_name):
            try:
                for line in iter(pipe.readline, ''):
                    output_queue.put((stream_name, line))
            except Exception as e:
                output_queue.put((stream_name, f"[Error reading stream: {e}]\n"))
            finally:
                pipe.close()
        threading.Thread(target=reader_thread, args=(process.stdout, 'stdout'), daemon=True).start()
        threading.Thread(target=reader_thread, args=(process.stderr, 'stderr'), daemon=True).start()
        try:
            return process.wait(timeout=timeout)
        except subprocess.TimeoutExpired:
            process.kill()
            output_queue.put(('stderr', "[Command timed out]\n"))
            return "TIMEOUT"
    except Exception as e:
        output_queue.put(('stderr', f"[ERROR] Could not start command: {e}\n"))
        return None

def run_pip_command(command, capture_output=True, text=True, check=False, timeout=15):
    pip_path = get_pip_path()
    if not pip_path:
        return "PIP_NOT_FOUND"
    
    # Create a clean environment without VIRTUAL_ENV
    env = os.environ.copy()
    if 'VIRTUAL_ENV' in env:
        del env['VIRTUAL_ENV']
    
    full_command = [pip_path] + command
    
    # Add startupinfo to prevent flickering CMD windows on Windows
    startupinfo = None
    if IS_WINDOWS:
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    
    try:
        return subprocess.run(
            full_command, capture_output=capture_output, text=text, check=check,
            encoding=locale.getpreferredencoding(False) or 'utf-8', errors='replace',
            env=env, timeout=timeout, shell=False,
            creationflags=CREATE_NO_WINDOW if IS_WINDOWS else 0,
            startupinfo=startupinfo  # Add startupinfo parameter
        )
    except Exception as e:
        print(f"Error in run_pip_command: {e}")
        return None

def get_package_summary(package_name):
    result = run_pip_command(["show", package_name], timeout=8)
    if not isinstance(result, subprocess.CompletedProcess) or result.returncode != 0:
        return f"Info not found ({result if not isinstance(result, subprocess.CompletedProcess) else result.stderr or result.returncode})."
    summary = "Summary not found."
    try:
        for line in result.stdout.splitlines():
            if line.startswith("Summary:"):
                summary = line.split(":", 1)[1].strip() or "No summary available."
                break
    except Exception:
        summary = "Error parsing package info."
    return summary

def get_installed_packages_fallback():
    """Use importlib.metadata to get installed packages, replacing pkg_resources."""
    if importlib_metadata is None:
        print("Error: importlib.metadata not available. Cannot list packages.")
        return []
    try:
        return [{'name': dist.name, 'version': dist.version} for dist in importlib_metadata.distributions()]
    except Exception as e:
        print(f"Fallback package listing failed: {e}")
        return []

def set_window_icon(window):
    """Set the window icon using platform-specific methods."""
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        
        if IS_WINDOWS:
            # Windows: Use .ico file
            icon_path = os.path.join(script_dir, "app_icon.ico")
            if os.path.exists(icon_path):
                try:
                    window.iconbitmap(icon_path)
                except Exception as e:
                    print(f"Windows iconbitmap failed: {e}")
        else:
            # macOS/Linux: Use .png file
            icon_path = os.path.join(script_dir, "app_icon_titlebar.png")
            if os.path.exists(icon_path):
                try:
                    photo = tk.PhotoImage(file=icon_path)
                    window.iconphoto(True, photo)
                except Exception as e:
                    print(f"iconphoto failed: {e}")
                
        print(f"Successfully attempted to set icon from: {icon_path}")
    except Exception as e:
        print(f"Error in set_window_icon: {e}")

def log_package_action(action, package_name, version=None):
    """Log package actions to the log file."""
    try:
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            if version:
                f.write(f"{timestamp} - {action}: {package_name}=={version}\n")
            else:
                # Try to get version from installed packages
                result = run_pip_command(["show", package_name], timeout=8)
                if isinstance(result, subprocess.CompletedProcess) and result.returncode == 0:
                    for line in result.stdout.splitlines():
                        if line.startswith("Version:"):
                            version = line.split(":", 1)[1].strip()
                            f.write(f"{timestamp} - {action}: {package_name}=={version}\n")
                            break
                    else:
                        f.write(f"{timestamp} - {action}: {package_name}\n")
                else:
                    f.write(f"{timestamp} - {action}: {package_name}\n")
    except Exception as e:
        print(f"Error writing to log file: {e}")

def log_app_event(event_type):
    """Log application events (start/stop)."""
    try:
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(f"{timestamp} - APPLICATION {event_type}\n")
    except Exception as e:
        print(f"Error writing to log file: {e}")

def count_application_starts():
    """Count the number of application starts in the log file."""
    try:
        if not os.path.exists(LOG_FILE):
            return 0
        count = 0
        with open(LOG_FILE, 'r', encoding='utf-8') as f:
            for line in f:
                if "APPLICATION STARTED" in line:
                    count += 1
        return count
    except Exception as e:
        print(f"Error counting application starts: {e}")
        return 0

def truncate_log_file():
    """Truncate the log file and add a header."""
    try:
        with open(LOG_FILE, 'w', encoding='utf-8') as f:
            f.write(f"=== Python Global Package Manager Log ===\n")
            f.write(f"Created: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"=======================================\n\n")
        return True
    except Exception as e:
        print(f"Error truncating log file: {e}")
        return False

def initialize_log_file():
    """Create log file if it doesn't exist and log application start."""
    try:
        # Ensure app data directory exists
        os.makedirs(APP_DATA_DIR, exist_ok=True)
        
        # Create log file if it doesn't exist
        if not os.path.exists(LOG_FILE):
            with open(LOG_FILE, 'w', encoding='utf-8') as f:
                f.write(f"=== Python Global Package Manager Log ===\n")
                f.write(f"Created: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"=======================================\n\n")
        
        # Check if we need to truncate (after 14 starts)
        if count_application_starts() >= 14:
            truncate_log_file()
        
        # Log application start
        log_app_event("STARTED")
    except Exception as e:
        print(f"Error initializing log file: {e}")

def clear_log_file():
    """Clear the log file."""
    try:
        with open(LOG_FILE, 'w', encoding='utf-8') as f:
            f.write("")
        return True
    except Exception as e:
        print(f"Error clearing log file: {e}")
        return False

def open_log_file():
    """Open the log file in the default text editor."""
    try:
        if IS_WINDOWS:
            os.startfile(LOG_FILE)
        elif IS_MAC:
            subprocess.run(['open', LOG_FILE])
        else:  # Linux
            subprocess.run(['xdg-open', LOG_FILE])
        return True
    except Exception as e:
        print(f"Error opening log file: {e}")
        return False

# When launching subprocesses, always use the venv Python path for venv-specific actions.
# venv_python = os.path.join(os.path.dirname(__file__), 'venv', 'Scripts', 'python.exe')
# Example: subprocess.run([venv_python, 'some_script.py'])

class PackageManagerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Initialize log file
        initialize_log_file()
        
        self.overrideredirect(True)
        self.drag_start_x = 0
        self.drag_start_y = 0
        self.tray_icon = None
        self.is_in_tray = False

        # Custom Title Bar
        self.title_bar = ctk.CTkFrame(self, fg_color=TITLE_BAR_COLOR, height=30, corner_radius=0)
        self.title_bar.pack(fill="x")
        
        # Add a placeholder for the icon in the title bar (pack it first, far left)
        self.title_icon_label = ctk.CTkLabel(self.title_bar, text="")
        self.title_icon_label.pack(side="left", padx=(5, 0))
        
        # Then add the title label
        self.title_label = ctk.CTkLabel(self.title_bar, text=APP_NAME, text_color="white", font=("Arial", 12, "bold"))
        self.title_label.pack(side="left", padx=10)
        # Add minimize and close buttons (right-aligned, flush with bar)
        self.close_button_title = ctk.CTkButton(self.title_bar, text="X", width=30, height=30, fg_color="red", hover_color="darkred", command=self.exit_application, corner_radius=0)
        self.close_button_title.pack(side="right", padx=(0,0), pady=0)
        self.minimize_button = ctk.CTkButton(self.title_bar, text="_", width=30, height=30, fg_color="#888", hover_color="#555", command=self.minimize_to_tray, corner_radius=0)
        self.minimize_button.pack(side="right", padx=(0,0), pady=0)
        # Add tooltips for buttons
        CTkToolTip(self.close_button_title, message="Exit the application.")
        CTkToolTip(self.minimize_button, message="Minimize to system tray.")
        self.title_bar.bind("<Button-1>", self.start_drag)
        self.title_label.bind("<Button-1>", self.start_drag)
        self.title_bar.bind("<B1-Motion>", self.on_drag)
        self.title_label.bind("<B1-Motion>", self.on_drag)

        # Set the icon for the main window only
        self.current_icon_path = self.load_icon_from_log()
        if self.current_icon_path and os.path.exists(self.current_icon_path):
            try:
                self.iconbitmap(self.current_icon_path)
                # Set title bar icon
                from PIL import Image
                png_icon_path = os.path.join(os.path.dirname(self.current_icon_path), "app_icon_titlebar.png")
                img = Image.open(self.current_icon_path)
                if hasattr(img, 'n_frames') and img.n_frames > 1:
                    max_size = (0, 0)
                    max_frame = 0
                    for i in range(img.n_frames):
                        img.seek(i)
                        if img.size > max_size:
                            max_size = img.size
                            max_frame = i
                    img.seek(max_frame)
                img = img.convert("RGBA")
                img = img.resize((20, 20), Image.LANCZOS)
                img.save(png_icon_path)
                icon_image = ctk.CTkImage(light_image=Image.open(png_icon_path), size=(20, 20))
                self.title_icon_label.configure(image=icon_image)
                self.title_icon_label.image = icon_image
            except Exception as e:
                print(f"Could not load icon: {e}")
                self.current_icon_path = None

        # Menu Bar & Menus
        self.menu_bar_frame = ctk.CTkFrame(self, fg_color="transparent", height=30)
        self.menu_bar_frame.pack(fill="x")
        
        # File Menu
        self.file_button = ctk.CTkButton(self.menu_bar_frame, text="File", width=60, height=25, fg_color=MENU_BG_COLOR, hover_color=MENU_ACTIVE_BG_COLOR, text_color=MENU_FG_COLOR, font=MENU_FONT, command=self.show_file_menu, corner_radius=0)
        self.file_button.pack(side="left", padx=(5,0), pady=2)
        self.file_menu = tk.Menu(self, tearoff=0, bg=MENU_BG_COLOR, fg=MENU_FG_COLOR, activebackground=MENU_ACTIVE_BG_COLOR, activeforeground=MENU_FG_COLOR, font=MENU_FONT)
        self.file_menu.add_command(label="Re-launch as Administrator", command=self.relaunch_as_admin)
        self.file_menu.add_command(label="Change Window Icon...", command=self.change_window_icon)
        self.file_menu.add_command(label="Python Script Delete", command=self.open_python_file_scanner)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Exit", command=self.exit_application)

        # Tools Menu
        self.tools_button = ctk.CTkButton(self.menu_bar_frame, text="Tools", width=60, height=25, fg_color=MENU_BG_COLOR, hover_color=MENU_ACTIVE_BG_COLOR, text_color=MENU_FG_COLOR, font=MENU_FONT, command=self.show_tools_menu, corner_radius=0)
        self.tools_button.pack(side="left", padx=(2,5), pady=2)
        self.tools_menu = tk.Menu(self, tearoff=0, bg=MENU_BG_COLOR, fg=MENU_FG_COLOR, activebackground=MENU_ACTIVE_BG_COLOR, activeforeground=MENU_FG_COLOR, font=MENU_FONT)
        
        # Create Log submenu
        self.log_menu = tk.Menu(self.tools_menu, tearoff=0, bg=MENU_BG_COLOR, fg=MENU_FG_COLOR, activebackground=MENU_ACTIVE_BG_COLOR, activeforeground=MENU_FG_COLOR, font=MENU_FONT)
        self.log_menu.add_command(label="View Log", command=self.open_log)
        self.log_menu.add_command(label="Clear Log", command=self.confirm_clear_log)
        
        # Add Log submenu to Tools menu
        self.tools_menu.add_cascade(label="Log", menu=self.log_menu)
        
        # Add Venv Creator to Tools menu
        self.tools_menu.add_command(label="Virtual Env and Script Association", command=self.open_venv_creator)

        # Help Menu
        self.help_button = ctk.CTkButton(self.menu_bar_frame, text="Help", width=60, height=25, fg_color=MENU_BG_COLOR, hover_color=MENU_ACTIVE_BG_COLOR, text_color=MENU_FG_COLOR, font=MENU_FONT, command=self.show_help_menu, corner_radius=0)
        self.help_button.pack(side="left", padx=(2,5), pady=2)
        self.help_menu = tk.Menu(self, tearoff=0, bg=MENU_BG_COLOR, fg=MENU_FG_COLOR, activebackground=MENU_ACTIVE_BG_COLOR, activeforeground=MENU_FG_COLOR, font=MENU_FONT)
        self.help_menu.add_command(label="Search for Package on PyPI", command=self.open_pypi_search)
        self.help_menu.add_separator()
        self.help_menu.add_command(label="About & File Locations", command=self.show_about_dialog)
        self.help_menu.add_command(label="How to Create EXE", command=self.show_exe_help_dialog)

        self.title(APP_NAME)
        self.geometry(GEOMETRY)
        
        self.packages_list = []
        self.displayed_packages = []
        self.selected_package_indices = []
        self.refresh_thread = None
        self.action_lock = threading.Lock()
        self.output_queue = queue.Queue()
        self.common_package_names = self.load_common_package_names()

        # Main Content Frame
        self.content_frame = ctk.CTkFrame(self)
        self.content_frame.pack(fill="both", expand=True)
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(1, weight=1)

        # Top Frame
        self.top_frame = ctk.CTkFrame(self.content_frame)
        self.top_frame.grid(row=0, column=0, padx=10, pady=(10,5), sticky="ew")
        self.top_frame.grid_columnconfigure((0,1,2,3), weight=1)
        self.refresh_button = ctk.CTkButton(self.top_frame, text="Refresh List", command=self.trigger_refresh)
        self.refresh_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        CTkToolTip(self.refresh_button, message="Reload package list.")
        self.search_entry = ctk.CTkEntry(self.top_frame, placeholder_text="Search packages...")
        self.search_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.search_entry.bind("<KeyRelease>", self.search_packages)
        CTkToolTip(self.search_entry, message="Filter packages by name.")
        self.export_button = ctk.CTkButton(self.top_frame, text="Export Current Package List", command=self.export_package_list)
        self.export_button.grid(row=0, column=2, padx=5, pady=5, sticky="ew")
        CTkToolTip(self.export_button, message="Save package list to file.")
        # Add Select All checkbox (top right)
        self.select_all_var = tk.BooleanVar()
        self.select_all_checkbox = ctk.CTkCheckBox(self.top_frame, text="Select All", variable=self.select_all_var, command=self.toggle_select_all)
        self.select_all_checkbox.grid(row=0, column=3, padx=5, pady=5, sticky="e")
        CTkToolTip(self.select_all_checkbox, message="Select or deselect all packages.")

        # Middle Frame
        self.middle_frame = ctk.CTkFrame(self.content_frame)
        self.middle_frame.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")
        self.middle_frame.grid_columnconfigure(0, weight=1)
        self.middle_frame.grid_rowconfigure(0, weight=1)
        self.package_list_scrollable_frame = ctk.CTkScrollableFrame(self.middle_frame, label_text="Installed Packages (Alphabetical)")
        self.package_list_scrollable_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        self.checkbox_vars = []  # List of (var, pkg_idx)
        self.checkbox_widgets = []  # List of CTkCheckBox widgets
        self.description_textbox = ctk.CTkTextbox(self.middle_frame, height=DESCRIPTION_BOX_HEIGHT, wrap="word", state="disabled", font=("Arial", 10))
        self.description_textbox.grid(row=1, column=0, padx=5, pady=(0,5), sticky="ew")
        CTkToolTip(self.description_textbox, message="Summary of selected package(s).")

        # Bottom Frame
        self.bottom_frame = ctk.CTkFrame(self.content_frame)
        self.bottom_frame.grid(row=2, column=0, padx=10, pady=(5,5), sticky="ew")
        self.bottom_frame.grid_columnconfigure((0,1,2,3), weight=1)
        self.install_button = ctk.CTkButton(self.bottom_frame, text="Install Package...", command=self.open_install_dialog_custom)
        self.install_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        CTkToolTip(self.install_button, message="Install a new package.")
        self.uninstall_button = ctk.CTkButton(self.bottom_frame, text="Uninstall Selected", command=self.uninstall_selected_package, fg_color="red", hover_color="darkred")
        self.uninstall_button.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        CTkToolTip(self.uninstall_button, message="Remove selected package(s).")
        self.update_python_button = ctk.CTkButton(self.bottom_frame, text="Update Selected", command=self.update_python_packages, fg_color="purple", hover_color="#5A005A")
        self.update_python_button.grid(row=0, column=2, padx=5, pady=5, sticky="ew")
        CTkToolTip(self.update_python_button, message="Update selected package(s).")
        self.close_button = ctk.CTkButton(self.bottom_frame, text="Close", command=self.exit_application)
        self.close_button.grid(row=0, column=3, padx=5, pady=5, sticky="ew")
        CTkToolTip(self.close_button, message="Exit the application.")

        # Terminal Output Area
        self.terminal_output = ctk.CTkTextbox(self.content_frame, height=(TERMINAL_FONT[1] * TERMINAL_HEIGHT + 10), font=TERMINAL_FONT, wrap="word", border_width=1, border_color="gray50", fg_color="#1D1F21", text_color="#C5C8C6", state="disabled")
        self.terminal_output.grid(row=3, column=0, padx=10, pady=(0,10), sticky="ew")
        self.terminal_output.tag_config("stdout", foreground="#81A2BE")
        self.terminal_output.tag_config("stderr", foreground="#CC6666")
        self.terminal_output.tag_config("info", foreground="#B5BD68")
        self.terminal_output.tag_config("error", foreground="#F0C674")
        self.create_terminal_context_menu()

        # Initial actions
        self.search_entry.delete(0, tk.END)
        self.trigger_refresh()
        self.process_output_queue()

        # Bind minimize event (for Alt+Space, taskbar minimize, etc.)
        self.protocol("WM_DELETE_WINDOW", self.exit_application)
        self.bind("<Unmap>", self._on_unmap)
        self.bind("<Map>", self._on_map)

    # --- Core Methods ---
    def relaunch_as_admin(self):
        if IS_WINDOWS:
            try:
                import ctypes
                self.update_terminal_output("[Info] Attempting to re-launch as administrator...\n", "info")
                script_path = os.path.abspath(sys.argv[0])
                params = " ".join([f'"{arg}"' for arg in sys.argv[1:]])
                ret = ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{script_path}" {params}', None, 1)
                if ret <= 32:
                    self.update_terminal_output(f"[Error] Failed to re-launch as admin (Code: {ret}).\n", "error")
                else:
                    self.update_terminal_output("[Info] Re-launch initiated. Exiting current instance...\n", "info")
                    self.exit_application()
            except Exception as e:
                self.update_terminal_output(f"[Error] Could not re-launch as admin: {e}\n", "error")
                self._show_ctk_message_dialog("Admin Re-launch Error", f"Error: {e}", dialog_type="error")
        else:
            self._show_ctk_message_dialog("Run as Admin", "On Linux/macOS, run script with 'sudo'.", dialog_type="info")

    def open_pypi_search(self):
        """Open PyPI search page in the default browser."""
        search_term = ctk.CTkInputDialog(text="Enter package name to search on PyPI:", title="PyPI Search").get_input()
        if search_term and search_term.strip():
            term = search_term.strip()
            url = f"https://pypi.org/search/?q={term}"
            webbrowser.open(url)
            self.update_terminal_output(f"[Info] Opened PyPI search for '{term}' in browser.\n", "info")
        elif search_term is not None:
            self._show_ctk_message_dialog("PyPI Search", "Search term cannot be empty.", dialog_type="warning")

    def show_file_menu(self):
        self.file_menu.post(self.file_button.winfo_rootx(), self.file_button.winfo_rooty() + self.file_button.winfo_height())

    def show_help_menu(self):
        self.help_menu.post(self.help_button.winfo_rootx(), self.help_button.winfo_rooty() + self.help_button.winfo_height())

    def start_drag(self, event):
        self.drag_start_x = event.x_root - self.winfo_x()
        self.drag_start_y = event.y_root - self.winfo_y()

    def on_drag(self, event):
        self.geometry(f"+{event.x_root - self.drag_start_x}+{event.y_root - self.drag_start_y}")

    def show_about_dialog(self):
        pip_path = get_pip_path() or "Not found"
        env_info = get_python_env_info()
        msg = f"{APP_NAME}\n\n"
        msg += f"Python: {env_info['executable']}\n"
        msg += f"Version: {env_info['version']}\n"
        msg += f"Platform: {env_info['platform']}\n"
        msg += f"Environment: {'Virtual' if env_info['is_venv'] else 'Global'}\n"
        if env_info['is_venv']:
            msg += f"Venv Path: {env_info['venv_path']}\n"
        msg += f"Pip: {pip_path}\n"
        msg += f"Data Dir: {APP_DATA_DIR}"
        self._show_ctk_message_dialog("About & File Locations", msg)

    def show_exe_help_dialog(self):
        msg = "Use PyInstaller: pyinstaller --onefile --windowed your_script.py"
        self._show_ctk_message_dialog("Creating an EXE", msg)

    def _show_ctk_message_dialog(self, title, message, dialog_type="info", icon_path=None):
        dialog = ctk.CTkToplevel(self)
        dialog.title(title)
        dialog.geometry("420x180")
        dialog.transient(self)
        dialog.grab_set()
        icon_path = icon_path or self.current_icon_path
        if icon_path and os.path.exists(icon_path):
            try:
                dialog.iconbitmap(icon_path)
            except Exception:
                pass
        color = {"info": "#1A73E8", "warning": "#F9A825", "error": "#D32F2F"}.get(dialog_type, "#1A73E8")
        ctk.CTkLabel(dialog, text=title, font=("Arial", 14, "bold"), text_color=color).pack(pady=(15, 5))
        ctk.CTkLabel(dialog, text=message, font=("Arial", 11), wraplength=380, justify="left").pack(pady=(0, 10))
        ctk.CTkButton(dialog, text="OK", width=90, command=dialog.destroy).pack(pady=10)
        dialog.wait_window()

    def load_common_package_names(self):
        # Ensure app data directory exists
        try:
            os.makedirs(APP_DATA_DIR, exist_ok=True)
        except Exception as e:
            print(f"Error creating app data directory: {e}")
            return []

        if not os.path.exists(PACKAGE_NAME_CACHE_FILE):
            dummy = ["requests", "numpy", "pandas", "matplotlib", "flask", "django", "customtkinter", "pillow", "packaging"]
            try:
                with open(PACKAGE_NAME_CACHE_FILE, 'w') as f:
                    json.dump(dummy, f)
            except Exception as e:
                print(f"Could not create dummy cache: {e}")
                return []
        try:
            with open(PACKAGE_NAME_CACHE_FILE, 'r') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading package cache: {e}")
            return []

    def search_packages(self, event=None):
        search_term = self.search_entry.get().strip().lower()
        self.displayed_packages = [pkg for pkg in self.packages_list if search_term in pkg['name'].lower()] if search_term else self.packages_list[:]
        self.display_packages()
        self.update_terminal_output(f"[Info] Filtered: {len(self.displayed_packages)} for '{search_term}'.\n", "info")

    def create_context_menu(self, listbox):
        menu = tk.Menu(listbox, tearoff=0)
        menu.add_command(label="Copy", command=lambda: self.copy_package_name(listbox))
        listbox.bind("<Button-3>", lambda event: menu.post(event.x_root, event.y_root))

    def create_terminal_context_menu(self):
        menu = tk.Menu(self.terminal_output, tearoff=0)
        menu.add_command(label="Copy", command=self.copy_terminal_output)
        self.terminal_output.bind("<Button-3>", lambda event: menu.post(event.x_root, event.y_root))

    def copy_package_name(self, listbox):
        sel_indices = listbox.curselection()
        sel_text = "\n".join([listbox.get(i) for i in sel_indices])
        if sel_text:
            self.clipboard_clear()
            self.clipboard_append(sel_text)
            self.update_terminal_output(f"[Info] Copied {len(sel_indices)} package(s).\n", "info")

    def copy_terminal_output(self):
        try:
            self.terminal_output.configure(state="normal")
            sel_text = self.terminal_output.selection_get()
            self.clipboard_clear()
            self.clipboard_append(sel_text)
            self.update_terminal_output("[Info] Copied terminal output.\n", "info")
        except tk.TclError:
            self.update_terminal_output("[Info] No text selected.\n", "info")
        finally:
            self.terminal_output.configure(state="disabled")

    def process_output_queue(self):
        try:
            while True:
                stream_name, line = self.output_queue.get_nowait()
                self.update_terminal_output(line, stream_name)
        except queue.Empty:
            pass
        self.after(100, self.process_output_queue)

    def update_terminal_output(self, text, tag="stdout"):
        self.terminal_output.configure(state="normal")
        self.terminal_output.insert(tk.END, text.rstrip('\n') + '\n', (tag,))
        self.terminal_output.see(tk.END)
        self.terminal_output.configure(state="disabled")

    def trigger_refresh(self, retry_count=0, max_retries=1):
        if self.refresh_thread and self.refresh_thread.is_alive():
            self.update_terminal_output("[Info] Refresh ongoing.\n", "info")
            return
        self.update_terminal_output("[Info] Fetching packages...\n", "info")
        self.disable_buttons()
        self.clear_list_area()
        self.clear_description()
        self.search_entry.delete(0, tk.END)
        self.displayed_packages = []
        self.refresh_thread = threading.Thread(target=self._get_and_display_packages_thread, args=(retry_count, max_retries), daemon=True)
        self.refresh_thread.start()

    def _get_and_display_packages_thread(self, retry_count, max_retries):
        success = self.fetch_package_list()
        if not success and retry_count < max_retries:
            self.after(0, lambda: self._retry_refresh(retry_count + 1, max_retries))
        else:
            self.after(0, self._update_ui_after_refresh, success)

    def _retry_refresh(self, retry_count, max_retries):
        self.update_terminal_output(f"[Info] Retry {retry_count}/{max_retries}...\n", "info")
        self.trigger_refresh(retry_count, max_retries)

    def fetch_package_list(self):
        result = run_pip_command(["list", "--format=json", "--disable-pip-version-check"], timeout=20)
        if not isinstance(result, subprocess.CompletedProcess) or result.returncode != 0 or not result.stdout:
            self.output_queue.put(('error', f"[Error] pip list failed. Trying fallback...\n"))
            self.packages_list = get_installed_packages_fallback()
            if not self.packages_list:
                self.output_queue.put(('error', "[Error] Fallback also failed.\n"))
                return False
            self.output_queue.put(('info', f"[Info] Loaded {len(self.packages_list)} via fallback.\n"))
            return True
        try:
            self.packages_list = [{'name': pkg['name'], 'version': pkg['version']} for pkg in json.loads(result.stdout)]
            return True
        except Exception as e:
            self.output_queue.put(('error', f"[Error] Parsing packages: {e}\n"))
            self.packages_list = []
            return False

    def _update_ui_after_refresh(self, success):
        if success:
            self.displayed_packages = self.packages_list[:]
            self.display_packages()
            self.update_terminal_output(f"[Info] Loaded {len(self.displayed_packages)} packages.\n", "info")
        else:
            self.clear_list_area()
            self.package_list_scrollable_frame.insert(tk.END, "Failed to load packages.")
        self.enable_buttons()

    def clear_list_area(self):
        # Remove all checkboxes from the scrollable frame
        for widget in self.package_list_scrollable_frame.winfo_children():
            widget.destroy()
        self.checkbox_vars = []
        self.checkbox_widgets = []
        self.selected_package_indices = []

    def clear_description(self):
        self.description_textbox.configure(state="normal")
        self.description_textbox.delete("1.0", tk.END)
        self.description_textbox.configure(state="disabled")

    def display_packages(self):
        self.clear_list_area()
        if not self.displayed_packages:
            ctk.CTkLabel(self.package_list_scrollable_frame, text="No packages to display.").grid(row=0, column=0, columnspan=NUM_COLUMNS, sticky="nsew")
            return
        self.displayed_packages.sort(key=lambda x: x['name'].lower())
        total_packages = len(self.displayed_packages)
        num_columns = NUM_COLUMNS
        pkg_per_col = math.ceil(total_packages / num_columns) if num_columns > 0 and total_packages > 0 else total_packages
        current_displayed_idx = 0
        for col_idx in range(num_columns):
            self.package_list_scrollable_frame.grid_columnconfigure(col_idx, weight=1)
            col_frame = ctk.CTkFrame(self.package_list_scrollable_frame, fg_color="transparent")
            col_frame.grid(row=0, column=col_idx, padx=2, pady=2, sticky="nsew")
            for row in range(pkg_per_col):
                if current_displayed_idx < total_packages:
                    pkg_info = self.displayed_packages[current_displayed_idx]
                    var = tk.BooleanVar()
                    cb = ctk.CTkCheckBox(col_frame, text=f"{pkg_info['name']}=={pkg_info['version']}", variable=var, command=self._on_checkbox_select)
                    cb.pack(anchor="w", pady=1, padx=2, fill="x")
                    self.checkbox_vars.append((var, current_displayed_idx))
                    self.checkbox_widgets.append(cb)
                    current_displayed_idx += 1
                else:
                    break
        # Update Select All checkbox state
        self.update_select_all_checkbox()

    def _on_checkbox_select(self):
        # Update selected_package_indices based on checked boxes
        self.selected_package_indices = [pkg_idx for var, pkg_idx in self.checkbox_vars if var.get()]
        self.update_description()
        self.update_select_all_checkbox()

    def update_select_all_checkbox(self):
        # Set the select all checkbox state based on current selection
        if not self.checkbox_vars:
            self.select_all_var.set(False)
            return
        all_selected = all(var.get() for var, _ in self.checkbox_vars)
        none_selected = all(not var.get() for var, _ in self.checkbox_vars)
        if all_selected:
            self.select_all_var.set(True)
        elif none_selected:
            self.select_all_var.set(False)
        else:
            self.select_all_var.set(False)

    def get_selected_package_info(self):
        if not self.displayed_packages or not self.selected_package_indices:
            return []
        infos = []
        for original_pkg_idx in self.selected_package_indices:
            if 0 <= original_pkg_idx < len(self.displayed_packages):
                infos.append(self.displayed_packages[original_pkg_idx])
        return sorted(infos, key=lambda p: p['name'].lower())

    def update_description(self):
        sel_pkg_info = self.get_selected_package_info()
        self.description_textbox.configure(state="normal")
        self.description_textbox.delete("1.0", tk.END)
        if sel_pkg_info:
            count = len(sel_pkg_info)
            self.description_textbox.insert("1.0", f"Selected {count} package(s):\n\n")
            for i, p_info in enumerate(sel_pkg_info):
                if i < 10 or count <= 10:
                    self.description_textbox.insert(tk.END, f"{p_info['name']}=={p_info['version']}\n")
                elif i == 10:
                    self.description_textbox.insert(tk.END, f"...and {count-10} more.\n")
                    break
        else:
            self.description_textbox.insert("1.0", "Select package(s) to see names.")
        self.description_textbox.configure(state="disabled")

    def _clear_selection_and_description(self):
        self.selected_package_indices = []
        for var, _ in self.checkbox_vars:
            var.set(False)
        self.update_description()

    # --- Action Methods ---
    def open_install_dialog_custom(self):
        InstallPackageDialog(master=self, title="Install Package", common_packages=self.common_package_names, app_instance=self, icon_path=self.current_icon_path)

    def _install_package_task(self, package_name):
        self.output_queue.put(('info', f"--- Installing {package_name} ---\n"))
        rc = run_pip_command_live(["install", package_name], self.output_queue)
        if rc == 0:
            # Get installed version
            result = run_pip_command(["show", package_name], timeout=8)
            version = None
            if isinstance(result, subprocess.CompletedProcess) and result.returncode == 0:
                for line in result.stdout.splitlines():
                    if line.startswith("Version:"):
                        version = line.split(":", 1)[1].strip()
                        break
            log_package_action("INSTALLED", package_name, version)
        self.output_queue.put(('info', f"--- Install {package_name} finished (Code: {rc}) ---\n"))
        return rc

    def uninstall_selected_package(self):
        sel_pkgs = self.get_selected_package_info()
        if not sel_pkgs:
            self._show_ctk_message_dialog("No Selection", "Select package(s) to uninstall.", icon_path=self.current_icon_path)
            return
        names = [p['name'] for p in sel_pkgs]
        export_uninstall = tk.BooleanVar(value=False)
        confirm = ctk.CTkToplevel(self)
        confirm.title("Confirm Uninstall")
        confirm.geometry("420x220")
        confirm.transient(self)
        confirm.grab_set()
        if self.current_icon_path and os.path.exists(self.current_icon_path):
            try:
                confirm.iconbitmap(self.current_icon_path)
            except Exception:
                pass
        ctk.CTkLabel(confirm, text=f"Uninstall these packages?", font=("Arial", 13, "bold")).pack(pady=(15, 5))
        ctk.CTkLabel(confirm, text=", ".join(names), font=("Arial", 11)).pack(pady=(0, 5))
        cb = ctk.CTkCheckBox(confirm, text="Export Package Uninstall List", variable=export_uninstall, font=("Arial", 10))
        cb.pack(pady=(0, 10))
        btn_frame = ctk.CTkFrame(confirm, fg_color="transparent")
        btn_frame.pack(pady=10)
        result = {'ok': False}
        def on_ok():
            result['ok'] = True
            confirm.destroy()
        def on_cancel():
            confirm.destroy()
        ctk.CTkButton(btn_frame, text="OK", width=90, command=on_ok).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="Cancel", width=90, command=on_cancel).pack(side="right", padx=10)
        confirm.wait_window()
        if not result['ok']:
            return
        if export_uninstall.get():
            fp = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text", "*.txt")], title="Save Uninstall List", initialfile="uninstall_list.txt")
            if not fp:
                return
            try:
                with open(fp, 'w') as f:
                    f.write("Uninstall Packages\n")
                    f.write(f"Date: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                    for pkg in sel_pkgs:
                        f.write(f"- {pkg['name']}=={pkg['version']}\n")
                self.update_terminal_output(f"[Info] Uninstall list exported to {fp}\n", "info")
            except Exception as e:
                self._show_ctk_message_dialog("Export Error", f"Failed to export uninstall list: {e}", icon_path=self.current_icon_path)
                return
        self.run_long_task(self._uninstall_package_task, names, "Uninstall")

    def _uninstall_package_task(self, pkg_names_list):
        overall_rc = 0
        for name in pkg_names_list:
            # Get version before uninstalling
            result = run_pip_command(["show", name], timeout=8)
            version = None
            if isinstance(result, subprocess.CompletedProcess) and result.returncode == 0:
                for line in result.stdout.splitlines():
                    if line.startswith("Version:"):
                        version = line.split(":", 1)[1].strip()
                        break
            
            self.output_queue.put(('info', f"--- Uninstalling {name} ---\n"))
            rc = run_pip_command_live(["uninstall", "-y", name], self.output_queue)
            if rc == 0:
                log_package_action("UNINSTALLED", name, version)
            self.output_queue.put(('info', f"--- Uninstall {name} finished (Code: {rc}) ---\n"))
            overall_rc = rc if rc != 0 and overall_rc == 0 else overall_rc
        if overall_rc == 0:
            self.after(0, self._clear_selection_and_description)
        return overall_rc

    def update_python_packages(self):
        sel_pkgs = self.get_selected_package_info()
        if not sel_pkgs:
            self._show_ctk_message_dialog("No Selection", "Select package(s) to update.", icon_path=self.current_icon_path)
            return
        names = [p['name'] for p in sel_pkgs]
        confirm = ctk.CTkToplevel(self)
        confirm.title("Confirm Update")
        confirm.geometry("420x180")
        confirm.transient(self)
        confirm.grab_set()
        if self.current_icon_path and os.path.exists(self.current_icon_path):
            try:
                confirm.iconbitmap(self.current_icon_path)
            except Exception:
                pass
        ctk.CTkLabel(confirm, text=f"Update these packages?", font=("Arial", 13, "bold")).pack(pady=(15, 5))
        ctk.CTkLabel(confirm, text=", ".join(names), font=("Arial", 11)).pack(pady=(0, 10))
        btn_frame = ctk.CTkFrame(confirm, fg_color="transparent")
        btn_frame.pack(pady=10)
        result = {'ok': False}
        def on_ok():
            result['ok'] = True
            confirm.destroy()
        def on_cancel():
            confirm.destroy()
        ctk.CTkButton(btn_frame, text="OK", width=90, command=on_ok).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="Cancel", width=90, command=on_cancel).pack(side="right", padx=10)
        confirm.wait_window()
        if not result['ok']:
            return
        self.run_long_task(self._update_python_task, names, "Update")

    def _update_python_task(self, pkg_names_list):
        overall_rc = 0
        for name in pkg_names_list:
            self.output_queue.put(('info', f"--- Updating {name} ---\n"))
            rc = run_pip_command_live(["install", "--upgrade", name], self.output_queue)
            if rc == 0:
                # Get new version after update
                result = run_pip_command(["show", name], timeout=8)
                version = None
                if isinstance(result, subprocess.CompletedProcess) and result.returncode == 0:
                    for line in result.stdout.splitlines():
                        if line.startswith("Version:"):
                            version = line.split(":", 1)[1].strip()
                            break
                log_package_action("UPDATED", name, version)
            self.output_queue.put(('info', f"--- Update {name} finished (Code: {rc}) ---\n"))
            overall_rc = rc if rc != 0 and overall_rc == 0 else overall_rc
        return overall_rc

    def export_package_list(self):
        if not self.displayed_packages:
            self._show_ctk_message_dialog("No Packages", "List is empty.")
            return
        fp = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text", "*.txt")], title="Save Package List", initialfile="requirements.txt")
        if not fp:
            return
        try:
            with open(fp, 'w') as f:
                for pkg in sorted(self.displayed_packages, key=lambda x: x['name'].lower()):
                    f.write(f"{pkg['name']}=={pkg['version']}\n")
            self.update_terminal_output(f"[Info] List exported to {fp}\n", "info")
            self._show_ctk_message_dialog("Export Successful", f"Saved to:\n{fp}")
        except Exception as e:
            self._show_ctk_message_dialog("Export Error", f"Failed: {e}")

    def run_long_task(self, task_func, task_args, action_name="Action"):
        if not self.action_lock.acquire(blocking=False):
            self._show_ctk_message_dialog("Busy", "Another operation is running.")
            return
        self.terminal_output.configure(state="normal")
        self.terminal_output.delete("1.0", tk.END)
        self.terminal_output.configure(state="disabled")
        self.update_terminal_output(f"[{action_name}] Starting for '{task_args}'...\n", "info")
        self.disable_buttons()
        threading.Thread(target=self._execute_task_and_update_ui, args=(task_func, task_args, action_name), daemon=True).start()

    def _execute_task_and_update_ui(self, task_func, task_args, action_name):
        status = "unknown"
        return_code = None
        try:
            return_code = task_func(task_args)
            if return_code == 0:
                status = "Success"
            elif isinstance(return_code, str):
                status = return_code
            else:
                status = "Failed" if return_code is not None else "Error (None returned)"
        except Exception as e:
            status = f"Error: {e}"
            self.output_queue.put(('error', f"[ERROR] Task {action_name}: {e}\n"))
        finally:
            def final_updates():
                msg = f"{action_name} for '{task_args}' finished.\nStatus: {status}"
                if status == "Success":
                    self._show_ctk_message_dialog(f"{action_name} Status", msg)
                else:
                    self._show_ctk_message_dialog(f"{action_name} Status", msg + "\nCheck terminal.")
                self.action_lock.release()
                self.after(500, self.trigger_refresh)
            if self.winfo_exists():
                self.after(0, final_updates)

    def disable_buttons(self):
        # Only disable buttons and checkboxes, not frames
        widgets = [self.refresh_button, self.install_button, self.uninstall_button, self.update_python_button, self.export_button]
        widgets += self.checkbox_widgets
        for w in widgets:
            if w and hasattr(w, 'configure') and w.winfo_exists():
                try:
                    w.configure(state="disabled")
                except Exception:
                    pass

    def enable_buttons(self):
        if self.action_lock.locked():
            return
        # Only enable buttons and checkboxes, not frames
        widgets = [self.refresh_button, self.install_button, self.uninstall_button, self.update_python_button, self.export_button]
        widgets += self.checkbox_widgets
        for w in widgets:
            if w and hasattr(w, 'configure') and w.winfo_exists():
                try:
                    w.configure(state="normal")
                except Exception:
                    pass

    def exit_application(self):
        """Save icon path and log application exit before closing."""
        self.save_icon_to_log()
        log_app_event("CLOSED")
        print("Exiting application...")
        self.destroy()

    def change_window_icon(self):
        """Allow user to select and set a new window icon."""
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            icon_path = filedialog.askopenfilename(
                title="Select Window Icon",
                filetypes=[("Icon files", "*.ico"), ("All files", "*.*")],
                initialdir=script_dir
            )
            if not icon_path:
                return
                
            # Copy the selected icon to the app directory
            new_icon_path = os.path.join(script_dir, "app_icon.ico")
            try:
                import shutil
                shutil.copy2(icon_path, new_icon_path)
                self.current_icon_path = new_icon_path
                self.update_terminal_output(f"[Info] Icon file copied to: {new_icon_path}\n", "info")
            except Exception as e:
                self.update_terminal_output(f"[Error] Failed to copy icon file: {e}\n", "error")
                return

            # Set the .ico for the window/taskbar
            try:
                self.iconbitmap(new_icon_path)
            except Exception as e:
                self.update_terminal_output(f"[Error] Failed to set window icon: {e}\n", "error")
                return

            # Convert .ico to .png for the title bar
            try:
                png_icon_path = os.path.join(script_dir, "app_icon_titlebar.png")
                img = Image.open(new_icon_path)
                if hasattr(img, 'n_frames') and img.n_frames > 1:
                    max_size = (0, 0)
                    max_frame = 0
                    for i in range(img.n_frames):
                        img.seek(i)
                        if img.size > max_size:
                            max_size = img.size
                            max_frame = i
                    img.seek(max_frame)
                img = img.convert("RGBA")
                img = img.resize((20, 20), Image.LANCZOS)
                img.save(png_icon_path)
                icon_image = ctk.CTkImage(light_image=Image.open(png_icon_path), size=(20, 20))
                self.title_icon_label.configure(image=icon_image)
                self.title_icon_label.image = icon_image
                self.update_terminal_output("[Info] Title bar icon updated successfully.\n", "info")
            except Exception as e:
                self.update_terminal_output(f"[Warning] Could not set title bar icon: {e}\n", "error")
        except Exception as e:
            self.update_terminal_output(f"[Error] Icon change failed: {e}\n", "error")
            self._show_ctk_message_dialog("Icon Error", f"Failed to change icon: {e}", dialog_type="error")

    def load_icon_from_log(self):
        """Load the icon path from the log file if it exists."""
        try:
            if os.path.exists(ICON_LOG_FILE):
                with open(ICON_LOG_FILE, 'r') as f:
                    path = f.read().strip()
                    if path and os.path.exists(path):
                        return path
        except Exception as e:
            print(f"Error loading icon from log: {e}")
        return None

    def save_icon_to_log(self):
        """Save the current icon path to the log file."""
        if self.current_icon_path:
            try:
                with open(ICON_LOG_FILE, 'w') as f:
                    f.write(self.current_icon_path)
            except Exception as e:
                print(f"Error saving icon path to log: {e}")

    def minimize_to_tray(self):
        """Minimize the application to system tray with platform-specific handling."""
        if not pystray or IS_MAC:  # Skip tray on macOS or if pystray is not available
            self.iconify()
            return
        
        if IS_LINUX:
            import importlib
            if not importlib.util.find_spec("Xlib"):
                print("python3-xlib not installed. System tray will not work on Linux.")
                self.iconify()
                return
            
        if not self.current_icon_path or not os.path.exists(self.current_icon_path):
            self.iconify()
            return
        
        self.withdraw()
        self.is_in_tray = True
        image = Image.open(self.current_icon_path)
        menu = pystray.Menu(pystray.MenuItem('Restore', self.restore_from_tray))
        self.tray_icon = pystray.Icon(APP_NAME, image, APP_NAME, menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def restore_from_tray(self, icon=None, item=None):
        self.is_in_tray = False
        if self.tray_icon:
            self.tray_icon.stop()
            self.tray_icon = None
        self.deiconify()
        self.lift()
        self.focus_force()

    def _on_unmap(self, event):
        # Handles minimize from taskbar or Alt+Space
        if not self.is_in_tray and self.state() == 'iconic':
            self.minimize_to_tray()

    def _on_map(self, event):
        # Handles restore from taskbar
        if not self.is_in_tray and self.state() == 'normal':
            self.deiconify()

    def toggle_select_all(self):
        value = self.select_all_var.get()
        for var, _ in self.checkbox_vars:
            var.set(value)
        self._on_checkbox_select()

    def show_tools_menu(self):
        self.tools_menu.post(self.tools_button.winfo_rootx(), self.tools_button.winfo_rooty() + self.tools_button.winfo_height())

    def open_log(self):
        if not os.path.exists(LOG_FILE):
            self._show_ctk_message_dialog("Log File", "Log file does not exist yet.", dialog_type="info")
            return
        if open_log_file():
            self.update_terminal_output("[Info] Opened log file.\n", "info")
        else:
            self._show_ctk_message_dialog("Error", "Could not open log file.", dialog_type="error")

    def confirm_clear_log(self):
        confirm = ctk.CTkToplevel(self)
        confirm.title("Confirm Clear Log")
        confirm.geometry("420x180")
        confirm.transient(self)
        confirm.grab_set()
        if self.current_icon_path and os.path.exists(self.current_icon_path):
            try:
                confirm.iconbitmap(self.current_icon_path)
            except Exception:
                pass
        ctk.CTkLabel(confirm, text="Are you sure you want to clear the log file?", font=("Arial", 13, "bold")).pack(pady=(15, 5))
        btn_frame = ctk.CTkFrame(confirm, fg_color="transparent")
        btn_frame.pack(pady=10, fill="x", expand=True)
        result = {'ok': False}
        def on_ok():
            result['ok'] = True
            confirm.destroy()
        def on_cancel():
            confirm.destroy()
        ctk.CTkButton(btn_frame, text="OK", width=90, command=on_ok).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="Cancel", width=90, command=on_cancel).pack(side="right", padx=10)
        confirm.wait_window()
        if result['ok']:
            if clear_log_file():
                self.update_terminal_output("[Info] Log file cleared.\n", "info")
            else:
                self._show_ctk_message_dialog("Error", "Could not clear log file.", dialog_type="error")

    def open_venv_creator(self):
        """Open the Virtual Environment Creator dialog."""
        # The VenvCreatorDialog now handles its own window management
        # with proper centering and visibility
        dialog = VenvCreatorDialog(master=self, icon_path=self.current_icon_path)
        # No need for additional setup - the dialog handles it internally

    def open_python_file_scanner(self):
        """Open the Python File Scanner in a new window."""
        try:
            script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "py_filescanner.py")
            if os.path.exists(script_path):
                self.update_terminal_output("[Info] Starting Python File Scanner...\n", "info")
                if IS_WINDOWS:
                    subprocess.Popen([sys.executable, script_path], creationflags=CREATE_NO_WINDOW)
                else:
                    subprocess.Popen([sys.executable, script_path])
            else:
                self.update_terminal_output("[Error] Could not find py_filescanner.py\n", "error")
                messagebox.showerror("File Not Found", "Could not find py_filescanner.py script")
        except Exception as e:
            self.update_terminal_output(f"[Error] Failed to launch Python File Scanner: {e}\n", "error")

class InstallPackageDialog(ctk.CTkToplevel):
    def __init__(self, master, title, common_packages, app_instance, icon_path=None):
        super().__init__(master)
        self.title(title)
        self.app_instance = app_instance
        self.common_package_names = common_packages or []
        self.icon_path = icon_path
        self.transient(master)
        self.grab_set()
        self.geometry("450x350")
        self.protocol("WM_DELETE_WINDOW", self._on_cancel_dialog)
        if self.icon_path and os.path.exists(self.icon_path):
            try:
                self.iconbitmap(self.icon_path)
            except Exception:
                pass
        ctk.CTkLabel(self, text="Enter package name or select below:").pack(pady=10)
        self.entry = ctk.CTkEntry(self, width=420)
        self.entry.pack(pady=5, padx=10)
        self.entry.bind("<KeyRelease>", self.update_search_results)
        self.entry.focus_set()
        self.results_listbox = Listbox(self, height=10, width=60, bg="#2B2B2B", fg="white", borderwidth=0, highlightthickness=0, exportselection=False)
        self.results_listbox.pack(pady=5, padx=10, fill="x", expand=True)
        self.results_listbox.bind("<Double-Button-1>", self.on_listbox_select_and_prepare_install)
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=10, fill="x", padx=10)
        self.install_btn = ctk.CTkButton(btn_frame, text="Install Package", command=self.on_install_button_pressed)
        self.install_btn.pack(side="left", expand=True, padx=5)
        self.cancel_btn = ctk.CTkButton(btn_frame, text="Cancel", command=self._on_cancel_dialog)
        self.cancel_btn.pack(side="right", padx=10, expand=True)
        self.update_search_results()

    def _on_cancel_dialog(self):
        self.destroy()

    def update_search_results(self, event=None):
        query = self.entry.get().strip()
        self.results_listbox.delete(0, tk.END)
        if not self.common_package_names:
            return
        display_list = []
        if not query:
            display_list = sorted(self.common_package_names)
        else:
            display_list = sorted([pkg for pkg in self.common_package_names if query.lower() in pkg.lower()])
        for item in display_list:
            self.results_listbox.insert(tk.END, item)

    def on_listbox_select_and_prepare_install(self, event=None):
        sel = self.results_listbox.curselection()
        if sel:
            selected_text = self.results_listbox.get(sel[0])
            pkg_name = selected_text.split(" (")[0].strip()
            self.entry.delete(0, tk.END)
            self.entry.insert(0, pkg_name)

    def on_install_button_pressed(self):
        pkg_name_original = self.entry.get().strip()
        if pkg_name_original:
            self.process_install_decision(pkg_name_original)
        else:
            if hasattr(self.app_instance, '_show_ctk_message_dialog'):
                self.app_instance._show_ctk_message_dialog("Input Error", "Package name cannot be empty.", dialog_type="warning")
            else:
                print("Input Error: Package name cannot be empty.")

    def process_install_decision(self, original_name):
        if not original_name:
            return
        normalized_name = canonicalize_name(original_name)
        name_to_install = original_name
        if original_name.lower().replace("_", "-").replace(".", "-") != normalized_name and original_name.lower() != normalized_name:
            confirm_dialog = ConfirmationDialog(master=self, title="Package Name Check",
                message_main=f"You entered: '{original_name}'", message_detail=f"This might be normalized to: '{normalized_name}' by pip.",
                option1_text=f"Install '{normalized_name}'", option2_text=f"Install '{original_name}' (as typed)", icon_path=self.icon_path)
            if hasattr(self.app_instance, 'wait_window'):
                self.app_instance.wait_window(confirm_dialog)
            else:
                self.wait_window(confirm_dialog)
            choice = confirm_dialog.choice
            if choice == "option1":
                name_to_install = normalized_name
            elif choice == "option2":
                name_to_install = original_name
            else:
                return
        self.app_instance.update_terminal_output(f"[Info] Proceeding to install: '{name_to_install}'\n", "info")
        self.app_instance.run_long_task(self.app_instance._install_package_task, name_to_install, action_name="Install")
        self.destroy()

class ConfirmationDialog(ctk.CTkToplevel):
    def __init__(self, master, title, message_main, message_detail, option1_text, option2_text, icon_path=None):
        super().__init__(master)
        self.title(title)
        self.icon_path = icon_path
        self.transient(master)
        self.grab_set()
        self.geometry("450x200")
        self.choice = None
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        if self.icon_path and os.path.exists(self.icon_path):
            try:
                self.iconbitmap(self.icon_path)
            except Exception:
                pass
        ctk.CTkLabel(self, text=message_main, font=("Arial", 14, "bold")).pack(pady=(10,2))
        ctk.CTkLabel(self, text=message_detail, font=("Arial", 12)).pack(pady=(0,10))
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=10, fill="x", expand=True)
        self.option1_btn = ctk.CTkButton(btn_frame, text=option1_text, command=self._on_option1)
        self.option1_btn.pack(side="left", padx=10, expand=True)
        self.option2_btn = ctk.CTkButton(btn_frame, text=option2_text, command=self._on_option2)
        self.option2_btn.pack(side="left", padx=10, expand=True)
        self.cancel_btn = ctk.CTkButton(btn_frame, text="Cancel", command=self._on_cancel, fg_color="gray50", hover_color="gray40")
        self.cancel_btn.pack(side="right", padx=10, expand=True)
        self.option1_btn.focus_set()

    def _on_option1(self):
        self.choice = "option1"
        self.destroy()

    def _on_option2(self):
        self.choice = "option2"
        self.destroy()

    def _on_cancel(self):
        self.choice = "cancel"
        self.destroy()

if __name__ == "__main__":
    if IS_WINDOWS:
        try:
            import ctypes
            app_id = f"MyCompany.{APP_NAME.replace(' ', '')}.GUI.1"
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
        except Exception:
            pass
    try:
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")
    except Exception as e:
        print(f"CTk theme error: {e}")
    app = PackageManagerApp()
    app.mainloop()