"""
Faculty Role Assigner — Windows Launcher
Auto-opens the browser and runs the web app on localhost.
"""
import sys
import os
import time
import socket
import threading
import webbrowser
import logging

# Silence Flask/Werkzeug console output
logging.getLogger("werkzeug").setLevel(logging.ERROR)

# When frozen by PyInstaller, add the bundle path so imports work
if getattr(sys, "frozen", False):
    os.chdir(os.path.dirname(sys.executable))
    sys.path.insert(0, sys._MEIPASS)

from app_prot import app

PORT = 5000

def find_free_port(start=5000):
    """Find an open port starting from `start`."""
    for port in range(start, start + 20):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    return start

def open_browser(port):
    time.sleep(1.5)
    webbrowser.open(f"http://localhost:{port}")

def main():
    port = find_free_port(PORT)

    print("=" * 54)
    print("  Faculty Role Assigner")
    print(f"  Open your browser at: http://localhost:{port}")
    print("  Close this window to stop.")
    print("=" * 54)

    threading.Thread(target=open_browser, args=(port,), daemon=True).start()

    app.run(host="127.0.0.1", port=port, debug=False, threaded=True, use_reloader=False)

if __name__ == "__main__":
    main()
