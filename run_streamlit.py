"""Run Streamlit while copying its live output to the console and log files."""

import os
from pathlib import Path
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
    logs = base / "logs" / "streamlit"
    logs.mkdir(parents=True, exist_ok=True)
    environment = os.environ.copy()
    environment.update(PYTHONUNBUFFERED="1", PYTHONIOENCODING="utf-8")
    for console in (sys.stdout, sys.stderr):
        console.reconfigure(encoding="utf-8", errors="replace")
    command = [
        sys.executable, "-u", "-m", "streamlit", "run", str(base / "main.py"),
        "--server.address=127.0.0.1", "--browser.serverAddress=localhost",
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
