"""Exercise the hidden CAD server through ready files, without WPS or HTTP."""
from pathlib import Path
import argparse
import ctypes
from ctypes import wintypes
import hashlib
import json
import os
import subprocess
import tempfile
import time
import uuid

ROOT = Path(__file__).resolve().parents[2]
REPO = ROOT.parent
TEMP = Path(tempfile.gettempdir()).resolve()
FLAGS = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
HEARTBEAT = None


def read_json(path):
    try:
        return json.loads(path.read_text(encoding="ascii"))
    except (OSError, ValueError):
        return None


def wait_for(predicate, description, timeout=45):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        result = predicate()
        if result:
            return result
        time.sleep(0.03)
    raise AssertionError("Timed out waiting for " + description)


class QueuedJob:
    def __init__(self, worker):
        self.worker = worker
        self.worker_hash = hashlib.sha256(worker.read_bytes()).hexdigest()
        self.id = uuid.uuid4().hex
        self.prefix = Path(os.environ["TEMP"]) / ("DongsidayeWpsCad_" + self.id)

    def path(self, phase, extension):
        return Path(str(self.prefix) + "." + phase + extension)

    def enqueue(self, phase, **options):
        request = self.path(phase, ".json")
        request.write_text(json.dumps(dict(id=self.id, phase=phase, **options), ensure_ascii=True), encoding="ascii")
        marker = self.path(phase, ".ready")
        marker.write_text(json.dumps({"exe": str(self.worker), "request": str(request), "workerHash": self.worker_hash}), encoding="ascii")
        return marker

    def finish(self, phase):
        def terminal():
            status = read_json(self.path(phase, ".status.json"))
            return status if status and status.get("state") == "done" else None
        status = wait_for(terminal, phase + " terminal status")
        assert status["ok"], status
        wait_for(lambda: not self.path(phase, ".ready").exists(), "consumed ready marker")
        return read_json(Path(status["output"])) if status.get("output") else None

    def cleanup(self):
        self.enqueue("cleanup")
        self.finish("cleanup")
        assert not self.prefix.exists()
        self.path("cleanup", ".status.json").unlink()


def test_cleanup_ack_before_reap(worker, server_pid):
    job = QueuedJob(worker)
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    native = ctypes.WinDLL("ntdll")
    kernel.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, wintypes.LPVOID, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    kernel.CreateFileW.restype = wintypes.HANDLE
    kernel.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel.OpenProcess.restype = wintypes.HANDLE
    kernel.CloseHandle.argtypes = [wintypes.HANDLE]
    native.NtSuspendProcess.argtypes = [wintypes.HANDLE]
    native.NtSuspendProcess.restype = ctypes.c_long
    native.NtResumeProcess.argtypes = [wintypes.HANDLE]
    native.NtResumeProcess.restype = ctypes.c_long
    held_lock = None
    server_handle = None
    suspended = False
    try:
        # Hold the job lock so the worker cannot finish before we pause our own
        # test server. The worker remains independent and can then finish while
        # the server is suspended, making the acknowledgement race deterministic.
        held_lock = kernel.CreateFileW(str(job.prefix) + ".lock", 0xC0000000, 1, None, 4, 0x80, None)
        assert held_lock != wintypes.HANDLE(-1).value, ctypes.get_last_error()
        marker = job.enqueue("cleanup")
        wait_for(lambda: not marker.exists(), "cleanup worker to start", 3)
        server_handle = kernel.OpenProcess(0x0800, False, server_pid)
        assert server_handle, ctypes.get_last_error()
        assert native.NtSuspendProcess(server_handle) == 0
        suspended = True
        kernel.CloseHandle(held_lock)
        held_lock = None
        status_path = job.path("cleanup", ".status.json")
        status = wait_for(lambda: read_json(status_path), "cleanup terminal status", 3)
        assert status["ok"] and status["state"] == "done", status
        wait_for(lambda: not Path(str(job.prefix) + ".lock").exists(), "worker final cleanup", 3)
        time.sleep(0.1)
        status_path.unlink()
        assert native.NtResumeProcess(server_handle) == 0
        suspended = False
        time.sleep(0.5)
        assert not status_path.exists(), "The server recreated a failed status after successful cleanup was acknowledged"
        print("PASS acknowledged cleanup is not republished as a false failure during reap", flush=True)
    finally:
        if suspended:
            native.NtResumeProcess(server_handle)
        if held_lock is not None:
            kernel.CloseHandle(held_lock)
        if server_handle:
            kernel.CloseHandle(server_handle)
        job.path("cleanup", ".status.json").unlink(missing_ok=True)


def main():
    global HEARTBEAT
    parser = argparse.ArgumentParser()
    parser.add_argument("--worker", type=Path, default=ROOT / "build/cad/WpsCadBridge.exe")
    options = parser.parse_args()
    worker = options.worker.resolve()
    assert worker.is_file()
    worker_hash = hashlib.sha256(worker.read_bytes()).hexdigest()
    HEARTBEAT = TEMP / ("DongsidayeWpsCad." + worker_hash + ".server.json")
    server = subprocess.Popen([str(worker), "--serve"], creationflags=FLAGS)
    owns_server = False
    jobs = []
    try:
        def live_server():
            heartbeat = read_json(HEARTBEAT)
            if not heartbeat or heartbeat.get("protocol") != 1:
                return None
            if Path(heartbeat["exe"]).resolve() != worker or abs(time.time() * 1000 - heartbeat["heartbeatAt"]) > 8000:
                return None
            return heartbeat
        heartbeat = wait_for(live_server, "CAD server heartbeat", 10)
        owns_server = heartbeat["pid"] == server.pid
        assert heartbeat["instance"] and heartbeat["startedTick"] > 0
        assert heartbeat["workerHash"] == worker_hash
        duplicate = subprocess.run([str(worker), "--serve"], timeout=10, creationflags=FLAGS)
        assert duplicate.returncode == 0
        assert read_json(HEARTBEAT)["pid"] == heartbeat["pid"], "A duplicate server replaced the live owner"
        print("PASS heartbeat protocol, physical worker path and per-user/exe singleton", flush=True)

        with tempfile.TemporaryDirectory(prefix="WpsCadServerFixture-") as work:
            fixture = Path(work) / "drawing"
            subprocess.run(["dotnet", "run", "--project", str(REPO / "ppt_addin/cad_converter/smoke/SmokeTest.csproj"), "-c", "Release", "--", str(fixture)], check=True, stdout=subprocess.PIPE, timeout=120, creationflags=FLAGS)
            drawing = next(fixture.glob("*.dwg"))
            job = QueuedJob(worker)
            jobs.append(job)
            job.enqueue("convert", input=str(drawing), ocr=False)
            time.sleep(0.02)
            # The second phase is queued before the first exits. It must wait
            # for that id instead of starting against an incomplete snapshot.
            job.enqueue("prepare", font="Microsoft YaHei", pageWidth=960, pageHeight=540, ocrSelections=[])
            converted = job.finish("convert")
            plan = job.finish("prepare")
            assert converted["textCount"] > 0 and plan["texts"]
            assert plan["texts"][0]["metrics"]["width"] > 0
            job.enqueue("finalize", fallbackIds=[])
            finalized = job.finish("finalize")
            assert Path(finalized["path"]).is_file()
            assert finalized["path"].startswith(str(job.prefix) + os.sep)
            print("PASS queued convert/prepare/finalize with same-id serialization and short TEMP spelling", flush=True)

            status_path = job.path("finalize", ".status.json")
            before = status_path.read_bytes()
            job.enqueue("finalize", fallbackIds=[])
            wait_for(lambda: not job.path("finalize", ".ready").exists(), "duplicate ready suppression")
            assert status_path.read_bytes() == before, "A completed phase was run again"
            print("PASS completed status suppresses duplicate phase execution", flush=True)

            invalid = QueuedJob(worker)
            bad_request = invalid.path("cleanup", ".json")
            bad_request.write_text(json.dumps({"id": invalid.id, "phase": "cleanup"}), encoding="ascii")
            bad_ready = invalid.path("cleanup", ".ready")
            bad_ready.write_text(json.dumps({"exe": str(worker.parent / "wrong-worker.exe"), "request": str(bad_request), "workerHash": worker_hash}), encoding="ascii")
            time.sleep(0.4)
            assert bad_ready.exists() and not invalid.path("cleanup", ".status.json").exists()
            bad_ready.unlink()
            bad_request.unlink()
            print("PASS mismatched worker markers cannot execute a job", flush=True)

            batch = [QueuedJob(worker) for _ in range(3)]
            jobs.extend(batch)
            for item in batch:
                item.enqueue("cleanup")
            for item in batch:
                item.finish("cleanup")
                item.path("cleanup", ".status.json").unlink()
            jobs = [job]
            print("PASS multiple independent ready jobs drain through the bounded queue", flush=True)
            job.cleanup()
            jobs.clear()
            if owns_server:
                test_cleanup_ack_before_reap(worker, server.pid)
            else:
                print("SKIP deterministic cleanup ack race requires a test-owned server", flush=True)
        old_timestamp = heartbeat["heartbeatAt"]
        wait_for(lambda: (read_json(HEARTBEAT) or {}).get("heartbeatAt", 0) > old_timestamp, "heartbeat refresh", 5)
        print("PASS heartbeat continues after jobs finish and owned cleanup completes", flush=True)
    finally:
        for job in jobs:
            try:
                job.cleanup()
            except Exception:
                pass
        if owns_server and server.poll() is None:
            server.terminate()
            server.wait(timeout=10)
            current = read_json(HEARTBEAT)
            if current and current.get("pid") == server.pid:
                HEARTBEAT.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
