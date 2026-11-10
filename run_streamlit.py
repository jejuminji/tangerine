"""Run Streamlit while copying its live output to the console and log files."""

import os
from datetime import datetime
from pathlib import Path
import socket
import subprocess
import sys
import threading


def _copy_output(source, console, logfile):
    for line in source:
        console.write(line)
        console.flush()
        logfile.write(line)
        logfile.flush()


def main() -> int:
    base = Path(__file__).resolve().parent
    for console in (sys.stdout, sys.stderr):
        console.reconfigure(encoding="utf-8", errors="replace")
    port = int(os.environ.get("STREAMLIT_SERVER_PORT", "8501"))
    with socket.socket() as probe:
        probe.settimeout(0.5)
        if probe.connect_ex(("127.0.0.1", port)) == 0:
            print(f"[ERROR] localhost:{port} is already in use. Tangerine may already be running.", flush=True)
            print("Close the existing Tangerine CMD window, or stop it with Ctrl+C, then run again.", flush=True)
            return 2
    logs = base / "logs" / "streamlit" / f"{datetime.now():%Y%m%d_%H%M%S}_{os.getpid()}"
    try:
        logs.mkdir(parents=True, exist_ok=False)
    except OSError as error:
        print(f"[ERROR] Cannot create the log folder: {error}", file=sys.stderr, flush=True)
        return 1
    print(f"Streamlit logs for this run: {logs}", flush=True)
    environment = os.environ.copy()
    environment.update(PYTHONUNBUFFERED="1", PYTHONIOENCODING="utf-8")
    command = [
        sys.executable, "-u", "-m", "streamlit", "run", str(base / "main.py"),
        "--server.address=127.0.0.1", "--browser.serverAddress=localhost",
        f"--server.port={port}",
    ]
    with (
        (logs / "streamlit.out.log").open("a", encoding="utf-8") as output_log,
        (logs / "streamlit.err.log").open("a", encoding="utf-8") as error_log,
    ):
        process = subprocess.Popen(
            command, cwd=base, env=environment, stdout=subprocess.PIPE,
            stderr=subprocess.PIPE, text=True, encoding="utf-8", errors="replace", bufsize=1,
        )
        pumps = [
            threading.Thread(target=_copy_output, args=(process.stdout, sys.stdout, output_log)),
            threading.Thread(target=_copy_output, args=(process.stderr, sys.stderr, error_log)),
        ]
        for pump in pumps:
            pump.start()
        try:
            return process.wait()
        except KeyboardInterrupt:
            if process.poll() is None:
                process.terminate()
            process.wait()
            return 130
        finally:
            for pump in pumps:
                pump.join()
            process.stdout.close()
            process.stderr.close()


if __name__ == "__main__":
    raise SystemExit(main())
