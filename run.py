import os
import sys
import time
import socket
import threading
import subprocess
import urllib.request
import urllib.error
from pathlib import Path


def get_app_dir():
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS)
    return Path(__file__).parent


def find_system_python():
    import winreg
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Python\PythonCore\3.14\InstallPath")
        python_dir, _ = winreg.QueryValueEx(key, "")
        python_exe = os.path.join(python_dir, "python.exe")
        if os.path.isfile(python_exe):
            return python_exe
    except Exception:
        pass

    for name in ["python.exe", "python3.exe"]:
        for path_dir in os.environ.get("PATH", "").split(os.pathsep):
            candidate = os.path.join(path_dir, name)
            if os.path.isfile(candidate):
                return candidate

    for candidate in [r"C:\Python314\python.exe", r"C:\Python3\python.exe"]:
        if os.path.isfile(candidate):
            return candidate

    return None


def wait_server_ready(host, port, timeout=30):
    url = f"http://{host}:{port}/_stcore/health"
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with urllib.request.urlopen(url, timeout=1) as resp:
                if resp.status == 200:
                    return True
        except (urllib.error.URLError, OSError, ConnectionError):
            pass
        time.sleep(0.5)
    return False


def main():
    app_dir = get_app_dir()
    script_path = str(app_dir / "streamlit_app.py")
    os.chdir(str(app_dir))

    HOST = "127.0.0.1"
    PORT = 8501

    python_exe = find_system_python()
    if not python_exe:
        from streamlit.web import bootstrap
        bootstrap.run(script_path, False, [], {
            "browser.gatherUsageStats": False,
            "server.headless": True,
        })
        return

    cmd = [
        python_exe, "-m", "streamlit", "run", script_path,
        "--server.port", str(PORT),
        "--server.headless", "true",
        "--browser.gatherUsageStats", "false",
    ]

    creation_flags = 0
    if sys.platform == "win32":
        creation_flags = subprocess.CREATE_NO_WINDOW

    proc = subprocess.Popen(cmd, cwd=str(app_dir), creationflags=creation_flags)

    opener_thread = threading.Thread(
        target=lambda: (
            wait_server_ready(HOST, PORT) and os.startfile(f"http://{HOST}:{PORT}")
        ),
        daemon=True,
    )
    opener_thread.start()

    try:
        proc.wait()
    except KeyboardInterrupt:
        proc.terminate()
        proc.wait()


if __name__ == "__main__":
    main()
