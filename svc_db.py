import os
import sys
import signal
import subprocess
import win32event
import win32service
import win32serviceutil


class DBApiService(win32serviceutil.ServiceFramework):
    _svc_name_ = "AudacesDBAPI"
    _svc_display_name_ = "Audaces DB API"
    _svc_description_ = "Servidor FastAPI (SQLite/Postgres) para Orcamentos"

    def __init__(self, args):
        super().__init__(args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.proc: subprocess.Popen | None = None

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        try:
            if self.proc and self.proc.poll() is None:
                self.proc.terminate()
        except Exception:
            pass
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        cwd = os.path.dirname(os.path.abspath(__file__))
        py = os.path.join(cwd, ".venv", "Scripts", "python.exe")
        if not os.path.exists(py):
            py = sys.executable
        cmd = [py, "-m", "uvicorn", "server_db:app", "--host", "0.0.0.0", "--port", "8000"]
        env = os.environ.copy()
        try:
            self.proc = subprocess.Popen(cmd, cwd=cwd, env=env)
            # Espera até receber sinal de stop
            win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE)
        finally:
            if self.proc and self.proc.poll() is None:
                try:
                    self.proc.terminate()
                except Exception:
                    pass


if __name__ == "__main__":
    win32serviceutil.HandleCommandLine(DBApiService)

