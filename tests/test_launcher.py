import ctypes
from ctypes import wintypes
import os
from pathlib import Path
import shutil
import socket
import subprocess
import sys
import tempfile
import time
import unittest
import urllib.request

import run_streamlit


class LauncherTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        shutil.copyfile(run_streamlit.__file__, self.root / "run_streamlit.py")
        (self.root / "main.py").write_text("import streamlit as st\nst.write('Ready')\n", encoding="utf-8")

    def environment(self, port):
        environment = os.environ.copy()
        environment.update(STREAMLIT_SERVER_PORT=str(port), STREAMLIT_SERVER_HEADLESS="true", STREAMLIT_BROWSER_GATHER_USAGE_STATS="false")
        return environment

    def test_occupied_port_is_reported_without_starting_another_server(self):
        with socket.socket() as listener:
            listener.bind(("127.0.0.1", 0))
            listener.listen()
            port = listener.getsockname()[1]
            result = subprocess.run(
                [sys.executable, str(self.root / "run_streamlit.py")],
                cwd=self.root, env=self.environment(port), capture_output=True,
                text=True, encoding="utf-8", timeout=10,
            )
        self.assertEqual(result.returncode, 2)
        self.assertIn("already in use", result.stdout)
        self.assertNotIn("Traceback", result.stderr)
        self.assertFalse((self.root / "logs").exists())

    @unittest.skipUnless(os.name == "nt", "Checks the Windows exclusive-file-lock regression")
    def test_old_log_locked_by_another_process_does_not_block_startup(self):
        logs = self.root / "logs" / "streamlit"
        logs.mkdir(parents=True)
        old_log = logs / "streamlit.out.log"
        old_log.write_text("Previous run\n", encoding="utf-8")
        create_file = ctypes.windll.kernel32.CreateFileW
        create_file.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
        create_file.restype = wintypes.HANDLE
        close_handle = ctypes.windll.kernel32.CloseHandle
        close_handle.argtypes = [wintypes.HANDLE]
        close_handle.restype = wintypes.BOOL
        handle = create_file(str(old_log), 0xC0000000, 0, None, 3, 128, None)
        self.assertNotEqual(handle, ctypes.c_void_p(-1).value)
        self.addCleanup(close_handle, handle)
        with self.assertRaises(PermissionError):
            old_log.open("a", encoding="utf-8")
        with socket.socket() as listener:
            listener.bind(("127.0.0.1", 0))
            port = listener.getsockname()[1]
        process = subprocess.Popen(
            [sys.executable, str(self.root / "run_streamlit.py")],
            cwd=self.root, env=self.environment(port), stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        )

        def stop():
            subprocess.run(["taskkill", "/PID", str(process.pid), "/T", "/F"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            process.communicate(timeout=10)

        self.addCleanup(stop)
        opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
        deadline = time.monotonic() + 25
        while time.monotonic() < deadline:
            self.assertIsNone(process.poll(), "Launcher exited before server startup")
            try:
                with opener.open(f"http://127.0.0.1:{port}/_stcore/health", timeout=1) as response:
                    self.assertEqual(response.status, 200)
                    break
            except OSError:
                time.sleep(0.1)
        else:
            self.fail("Streamlit did not start while the old log was locked")
        sessions = [path for path in logs.iterdir() if path.is_dir()]
        self.assertEqual(len(sessions), 1)
        self.assertTrue((sessions[0] / "streamlit.out.log").exists())
        self.assertTrue((sessions[0] / "streamlit.err.log").exists())


if __name__ == "__main__":
    unittest.main()
