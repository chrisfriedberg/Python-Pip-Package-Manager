
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

Tested on **Windows 10/11**, Python 3.10+

```bash


pip install -r requirements.txt
Main dependencies:
PySide6
customtkinter
pystray
Pillow
requests
pywin32
