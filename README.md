# 🧰 Python Package Manager App (v1.19)

A clean, modern, GUI-based package manager for Python, built with PySide6 and CustomTkinter. Designed to help developers, analysts, and support staff manage Python packages without touching the terminal.

---

## 🚀 Features

- ✅ Multi-select checkboxes for batch installs/uninstalls
- 🔍 Live search with real-time filter on installed packages
- 📦 Pip-backed install, uninstall, and upgrade logic
- 🌐 "Help" button opens selected package's PyPI page
- 📜 Logs pip output with timestamped entries
- 💾 Export current package list to a text file
- 🖼️ Tray icon support (with `.ico` fallback handling)
- 🔐 Auto-prompts for admin relaunch if elevation is needed

---

## 🛠 Requirements

Tested on **Windows 10/11**, **macOS**, and **Linux**, Python 3.10+

```bash
pip install -r requirements.txt
```

Main dependencies:
- customtkinter
- pystray
- Pillow
- requests
- packaging

---

## 📦 Installation

1. Clone or download this repository
2. Place the script in any directory (no specific location required)
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
4. Run the script:
   ```bash
   python package_manager_app.py
   ```

### Platform-Specific Notes

#### Windows
- Full system tray support
- Admin elevation prompts when needed

#### macOS
- System tray is disabled (macOS limitation)
- Uses standard window minimize/restore
- Requires Python 3.10+ with Tkinter support

#### Linux
- System tray requires `python3-xlib` package
- May require additional X11 dependencies

---

## 🔧 Troubleshooting

### Common Issues

1. **Script not found**: Ensure you're running the script from the correct directory or provide the full path
2. **Missing dependencies**: Run `pip install -r requirements.txt` to install all required packages
3. **macOS tray issues**: System tray is intentionally disabled on macOS due to platform limitations
4. **Linux tray issues**: Install `python3-xlib` package for system tray support

### File Locations

The app stores its data in platform-specific locations:
- Windows: `%APPDATA%\Python_Global_Package_Manager`
- macOS: `~/Library/Application Support/Python_Global_Package_Manager`
- Linux: `~/.local/share/Python_Global_Package_Manager`

---

## 📝 License

MIT License - See LICENSE file for details
