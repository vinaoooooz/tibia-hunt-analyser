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
    try:
        import winreg
        for ver in ["3.14", "3.13", "3.12", "3.11", "3.10"]:
            try:
                key = winreg.OpenKey(
                    winreg.HKEY_LOCAL_MACHINE,
                    rf"SOFTWARE\Python\PythonCore\{ver}\InstallPath",
                )
                python_dir, _ = winreg.QueryValueEx(key, "")
                python_exe = os.path.join(python_dir, "python.exe")
                if os.path.isfile(python_exe):
                    return python_exe
            except Exception:
                pass

        for hive in [winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE]:
            try:
                key = winreg.OpenKey(hive, r"SOFTWARE\Python\PythonCore")
                i = 0
                while True:
                    try:
                        subkey_name = winreg.EnumKey(key, i)
                        subkey = winreg.OpenKey(key, subkey_name + r"\InstallPath")
                        python_dir, _ = winreg.QueryValueEx(subkey, "")
                        python_exe = os.path.join(python_dir, "python.exe")
                        if os.path.isfile(python_exe):
                            return python_exe
                    except OSError:
                        break
                    i += 1
            except Exception:
                pass
    except ImportError:
        pass

    for name in ["python.exe", "python3.exe", "python314.exe", "python313.exe", "python312.exe"]:
        for path_dir in os.environ.get("PATH", "").split(os.pathsep):
            candidate = os.path.join(path_dir, name)
            if os.path.isfile(candidate):
                return candidate

    for ver in ["314", "313", "312", "311", "310"]:
        for base in [
            os.path.expandvars(r"%LOCALAPPDATA%\Programs\Python\Python3" + ver[-2:]),
            rf"C:\Python{ver}",
            rf"C:\Python3",
        ]:
            candidate = os.path.join(base, "python.exe")
            if os.path.isfile(candidate):
                return candidate

    for base in [
        os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\WindowsApps"),
        os.path.expandvars(r"%PROGRAMFILES%\WindowsApps"),
    ]:
        for name in ["python3.exe", "python.exe"]:
            candidate = os.path.join(base, name)
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


def open_browser(host, port):
    time.sleep(2)
    url = f"http://{host}:{port}"
    try:
        os.startfile(url)
    except Exception:
        import webbrowser
        webbrowser.open(url)


def show_error(message):
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Tibia Hunt Analyzer", message)
        root.destroy()
    except Exception:
        pass


def run_with_bootstrap(script_path):
    from streamlit.web import bootstrap

    HOST = "127.0.0.1"
    PORT = 8501

    opener = threading.Thread(
        target=lambda: open_browser(HOST, PORT),
        daemon=True,
    )
    opener.start()

    bootstrap.run(script_path, False, [], {
        "browser.gatherUsageStats": False,
        "server.headless": True,
    })


def run_with_subprocess(python_exe, script_path, app_dir):
    HOST = "127.0.0.1"
    PORT = 8501

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

    opener = threading.Thread(
        target=lambda: (
            wait_server_ready(HOST, PORT) and open_browser(HOST, PORT)
        ),
        daemon=True,
    )
    opener.start()

    try:
        proc.wait()
    except KeyboardInterrupt:
        proc.terminate()
        proc.wait()


def main():
    app_dir = get_app_dir()
    script_path = str(app_dir / "streamlit_app.py")
    os.chdir(str(app_dir))

    python_exe = find_system_python()
    if python_exe:
        run_with_subprocess(python_exe, script_path, app_dir)
    else:
        run_with_bootstrap(script_path)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        show_error(f"Erro ao iniciar o aplicativo:\n\n{e}")
