import os
import sys
import streamlit.web.cli as stcli

def resolve_path(path):
    # PyInstaller creates a temp folder and stores path in _MEIPASS
    if getattr(sys, 'frozen', False):
        resolved_path = os.path.abspath(os.path.join(sys._MEIPASS, path))
    else:
        resolved_path = os.path.abspath(os.path.join(os.getcwd(), path))
    return resolved_path

if __name__ == "__main__":
    # Ensure Playwright finds local browsers if frozen
    if getattr(sys, 'frozen', False):
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.abspath(os.path.join(os.path.dirname(sys.executable), "pw-browsers"))
        
    sys.argv = [
        "streamlit",
        "run",
        resolve_path("app.py"),
        "--global.developmentMode=false",
    ]
    sys.exit(stcli.main())
