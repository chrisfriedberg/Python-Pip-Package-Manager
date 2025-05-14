#!/usr/bin/env python
"""
Launcher script for running Python applications from a virtual environment.
Provides a system tray icon with exit functionality.
"""
import sys
import os
import subprocess
from PySide6.QtWidgets import QApplication, QSystemTrayIcon, QMenu
from PySide6.QtGui import QIcon, QAction

# --- VENV SAFETY CHECK ---
expected_venv = os.path.join(os.path.dirname(__file__), 'venv')
current_prefix = sys.prefix
if not os.path.normcase(current_prefix).startswith(os.path.normcase(expected_venv)):
    msg = f"ERROR: This launcher must be run from the venv at {expected_venv}.\nCurrent sys.prefix: {current_prefix}\nPlease use the venv's python.exe or pythonw.exe."
    print(msg)
    try:
        with open(os.path.join(os.path.dirname(__file__), 'launcher_error.log'), 'w') as f:
            f.write(msg)
    except Exception:
        pass
    sys.exit(1)

# Get the script directory
script_dir = os.path.dirname(os.path.abspath(__file__))

# Read association file
config = {}
assoc_file = os.path.join(script_dir, ".venv-association")
if os.path.exists(assoc_file):
    with open(assoc_file, "r") as f:
        for line in f:
            if "=" in line:
                key, value = line.strip().split("=", 1)
                config[key] = value

# Get paths from config
venv_rel_path = config.get("venv_path", "venv")
script_name = config.get("main_script", "")

# Build absolute paths
venv_path = os.path.join(script_dir, venv_rel_path)
script_path = os.path.join(script_dir, script_name)
venv_python = os.path.join(venv_path, "Scripts", "pythonw.exe")

# Import PySide6
venv_site_packages = os.path.join(venv_path, "Lib", "site-packages")
if venv_site_packages not in sys.path:
    sys.path.insert(0, venv_site_packages)

try:
    from PySide6.QtWidgets import QApplication, QSystemTrayIcon, QMenu, QAction
    from PySide6.QtGui import QIcon, QPixmap, QPainter, QColor, QBrush, QPen
    from PySide6.QtCore import QTimer
except ImportError:
    # Try to install PySide6 in the venv
    pip_path = os.path.join(venv_path, "Scripts", "pip.exe")
    if os.path.exists(pip_path):
        subprocess.run([pip_path, "install", "PySide6"], check=True)
        from PySide6.QtWidgets import QApplication, QSystemTrayIcon, QMenu, QAction
        from PySide6.QtGui import QIcon, QPixmap, QPainter, QColor, QBrush, QPen
        from PySide6.QtCore import QTimer

# Create function to generate an icon
def create_icon():
    """Create a simple icon for the system tray"""
    pixmap = QPixmap(64, 64)
    pixmap.fill(QColor(0, 0, 0, 0))
    painter = QPainter(pixmap)
    painter.setBrush(QBrush(QColor(0, 97, 255)))
    painter.setPen(QPen(QColor(0, 0, 0, 0)))
    painter.drawRect(0, 0, 64, 64)
    painter.setBrush(QBrush(QColor(255, 255, 255)))
    painter.drawRect(20, 20, 24, 24)
    painter.end()
    return pixmap

# Create QApplication instance
app = QApplication(sys.argv)
app.setQuitOnLastWindowClosed(False)

# Track all timers for cleanup
process_check_timer = None

# Full path to icon, or create one if it doesn't exist
icon_path = os.path.join(script_dir, "launcher_icon.ico")
icon = None
if os.path.exists(icon_path):
    icon = QIcon(icon_path)
else:
    # Create and save a default icon
    pixmap = create_icon()
    pixmap.save(icon_path)
    icon = QIcon(pixmap)

# Create tray icon with proper error handling
tray = QSystemTrayIcon()
tray.setIcon(icon)
tray.setToolTip(f"Launcher for {script_name}")

# Create menu for the tray icon
menu = QMenu()

# Process handle
process = None

# Function to properly clean up resources
def cleanup():
    print("[INFO] Cleaning up resources...")
    
    # Stop and delete the timer
    global process_check_timer
    if process_check_timer is not None:
        try:
            process_check_timer.stop()
            # In Qt, deleteLater is preferred over direct deletion
            process_check_timer.deleteLater()
            process_check_timer = None
            print("[INFO] Timer stopped and scheduled for deletion")
        except Exception as e:
            print(f"[ERROR] Failed to stop timer: {e}")
    
    # Hide tray before anything else
    if tray:
        try:
            tray.hide()
            print("[INFO] Tray icon hidden")
        except Exception as e:
            print(f"[ERROR] Failed to hide tray: {e}")
    
    # Kill any running process
    try:
        if process and process.poll() is None:
            process.terminate()
            print(f"[INFO] Process {process.pid} terminated")
    except Exception as e:
        print(f"[ERROR] Failed to terminate process: {e}")

# Hook cleanup to application exit
app.aboutToQuit.connect(cleanup)

# Launch the main script
try:
    # Use Popen to launch the script without waiting
    process = subprocess.Popen(
        [venv_python, script_path], 
        cwd=script_dir,
        creationflags=subprocess.CREATE_NO_WINDOW
    )
    print(f"[INFO] Launched process with PID: {process.pid}")
except Exception as e:
    error_msg = f"Error launching script: {str(e)}"
    print(f"[ERROR] {error_msg}")
    error_action = QAction(error_msg)
    error_action.setEnabled(False)
    menu.addAction(error_action)

def kill_app():
    """Terminate the process and exit the launcher"""
    print("[INFO] Exiting application...")
    cleanup()  # Ensure cleanup happens before quit
    QApplication.quit()

# Function to safely handle window actions
def handle_window_action(reason):
    # First check if process is running
    if process and process.poll() is None:
        # Process is running, but we don't know if window is accessible
        print("[INFO] Tray icon clicked, process is still running")
    else:
        # Process not running, exit
        print("[INFO] Tray icon clicked but main process not running, exiting")
        kill_app()

# Add tray icon click handler with window existence check
tray.activated.connect(handle_window_action)

# Periodically check if the main process is still alive
def check_process():
    if not process:
        print("[INFO] No process to monitor")
        return
        
    if process.poll() is not None:
        # Process has ended, clean up and exit
        print(f"[INFO] Main process ended with code {process.poll()}, exiting tray app")
        kill_app()

# Create and start the timer
process_check_timer = QTimer()
process_check_timer.timeout.connect(check_process)
process_check_timer.start(1000)  # Check every second
print("[INFO] Process monitor timer started")

# Add Exit option to the tray menu
exit_action = QAction("Exit App")
exit_action.triggered.connect(kill_app)
menu.addAction(exit_action)

# Display the tray icon
tray.setContextMenu(menu)
tray.show()
print("[INFO] Tray icon displayed")

# Run the event loop
print("[INFO] Entering main event loop")
sys.exit(app.exec()) 