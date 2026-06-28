import os
import subprocess
import sys

def main():
    print(">>> 1a. Generating PyInstaller spec file...")
    
    spec_content = """# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.utils.hooks import collect_all
from PyInstaller.utils.hooks import copy_metadata
import sys
sys.setrecursionlimit(5000)

datas = [('app.py', '.'), ('fireant_client.py', '.'), ('fireant_mapper.py', '.'), ('excel_writer.py', '.'), ('excel_inspector.py', '.'), ('html_parser.py', '.'), ('validators.py', '.')]
binaries = []
hiddenimports = ['openpyxl', 'pandas', 'bs4']
datas += collect_data_files('playwright')
datas += copy_metadata('streamlit')
tmp_ret = collect_all('streamlit')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['run_main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='CafeFTool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='CafeFTool',
)
"""
    
    with open("CafeFTool.spec", "w", encoding="utf-8") as f:
        f.write(spec_content)
        
    print(">>> 1c. Building from spec file...")
    subprocess.run([sys.executable, "-m", "PyInstaller", "--clean", "--noconfirm", "CafeFTool.spec"], check=True)
    
    print(">>> 2. Downloading Playwright Chromium browser...")
    target_pw_dir = os.path.abspath(os.path.join("dist", "CafeFTool", "pw-browsers"))
    
    # Create directory if not exists
    os.makedirs(target_pw_dir, exist_ok=True)
    
    # Force Playwright to install Chromium directly into our build folder
    env = os.environ.copy()
    env["PLAYWRIGHT_BROWSERS_PATH"] = target_pw_dir
    
    subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], env=env, check=True)
    
    print("\n>>> SUCCESS! The portable folder is located at 'dist/CafeFTool'.")
    print(">>> You can zip the 'CafeFTool' folder and send it to your friend.")

if __name__ == "__main__":
    main()
