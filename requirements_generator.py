import os
import subprocess
import sys
import re

# Step 1: Run pipreqs to generate requirements.txt
def run_pipreqs(target_dir):
    try:
        result = subprocess.run([sys.executable, '-m', 'pip', 'install', 'pipreqs'], capture_output=True, text=True)
        print(result.stdout)
        reqs_cmd = [sys.executable, '-m', 'pipreqs', target_dir, '--force']
        result = subprocess.run(reqs_cmd, capture_output=True, text=True)
        print(result.stdout)
        if result.returncode != 0:
            print("pipreqs failed:", result.stderr)
            return False
        return True
    except Exception as e:
        print(f"Error running pipreqs: {e}")
        return False

# Step 2: AI-powered check for import/package name mismatches and common mistakes
def analyze_requirements(target_dir):
    req_path = os.path.join(target_dir, 'requirements.txt')
    if not os.path.exists(req_path):
        print("requirements.txt not found!")
        return
    with open(req_path, 'r', encoding='utf-8') as f:
        lines = [line.strip() for line in f if line.strip() and not line.startswith('#')]

    # Known import-to-pypi mismatches
    known_map = {
        'PIL': 'pillow',
        'cv2': 'opencv-python',
        'sklearn': 'scikit-learn',
        'Crypto': 'pycryptodome',
        'yaml': 'pyyaml',
        'bs4': 'beautifulsoup4',
        'Image': 'pillow',
        'matplotlib.pyplot': 'matplotlib',
        'dateutil': 'python-dateutil',
        'lxml': 'lxml',
        'setuptools': 'setuptools',
        'wheel': 'wheel',
        'scipy': 'scipy',
        'pandas': 'pandas',
        'numpy': 'numpy',
        'requests': 'requests',
        'flask': 'flask',
        'django': 'django',
        'pytest': 'pytest',
        'tqdm': 'tqdm',
        'jupyter': 'notebook',
        'IPython': 'ipython',
        'openpyxl': 'openpyxl',
        'psycopg2': 'psycopg2-binary',
        'mysql': 'mysql-connector-python',
        'mysqlclient': 'mysqlclient',
        'sqlite3': None,  # stdlib
        'tkinter': None,  # stdlib
        'shutil': None,   # stdlib
        'os': None,       # stdlib
        'sys': None,      # stdlib
    }

    suggestions = []
    for pkg in lines:
        # Check for known mismatches
        base = re.split(r'[<>=]', pkg)[0].strip()
        if base in known_map:
            mapped = known_map[base]
            if mapped is None:
                suggestions.append(f"'{base}' is a standard library module and should not be in requirements.txt.")
            elif mapped != base:
                suggestions.append(f"'{base}' should be '{mapped}' in requirements.txt.")
        # Check for likely typos (very basic)
        elif base.lower() in known_map:
            mapped = known_map[base.lower()]
            if mapped:
                suggestions.append(f"'{base}' might be a typo. Did you mean '{mapped}'?")
        elif base.lower() == 'pyhton':
            suggestions.append("'pyhton' is a typo. Did you mean 'python'? (But python should not be in requirements.txt)")
        elif base.lower() == 'reqeusts':
            suggestions.append("'reqeusts' is a typo. Did you mean 'requests'?")
        # Add more AI-style heuristics here as needed

    print("\n=== AI-powered requirements.txt review ===")
    if suggestions:
        for s in suggestions:
            print("-", s)
    else:
        print("No obvious issues found. Your requirements.txt looks good!")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Generate and review requirements.txt using pipreqs and AI-powered checks.")
    parser.add_argument('target_dir', help="Path to your project directory")
    args = parser.parse_args()
    if run_pipreqs(args.target_dir):
        analyze_requirements(args.target_dir) 