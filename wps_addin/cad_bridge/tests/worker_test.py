"""Executable tests for the isolated WPS CAD worker and owned job protocol."""
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
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
REPO = ROOT.parent
TEMP = Path(tempfile.gettempdir()).resolve()
FLAGS = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


class Job:
    def __init__(self, worker):
        self.worker = worker
        self.id = uuid.uuid4().hex
        self.prefix = TEMP / ("DongsidayeWpsCad_" + self.id)
        self.folder = self.prefix

    def request(self, phase, **options):
        request = Path(str(self.prefix) + "." + phase + ".json")
        request.write_text(json.dumps(dict(id=self.id, phase=phase, **options), ensure_ascii=True), encoding="ascii")
        return request

    def status(self, phase):
        return Path(str(self.prefix) + "." + phase + ".status.json")

    def run(self, phase, expect=True, **options):
        request = self.request(phase, **options)
        process = subprocess.run([str(self.worker), "--job", str(request)], timeout=200, creationflags=FLAGS)
        status = json.loads(self.status(phase).read_text(encoding="utf-8"))
        assert status["state"] == "done", status
        assert status["ok"] is expect, status
        assert (process.returncode == 0) is expect, (process.returncode, status)
        output = json.loads(Path(status["output"]).read_text(encoding="ascii")) if status.get("output") else None
        return status, output

    def cleanup(self):
        self.run("cleanup")
        assert not self.folder.exists()
        leftovers = sorted(TEMP.glob(self.prefix.name + ".*"))
        assert [entry.resolve() for entry in leftovers] == [self.status("cleanup").resolve()], leftovers
        self.status("cleanup").unlink()


def prepare(job, **options):
    return job.run("prepare", font="Microsoft YaHei", pageWidth=960, pageHeight=540, ocrSelections=[], **options)


def test_real_conversion(worker, drawing):
    job = Job(worker)
    original = digest(drawing)
    try:
        _, converted = job.run("convert", input=str(drawing), ocr=False)
        assert converted["entityCount"] == 6
        assert converted["textCount"] >= 1
        assert converted["outputHash"] == digest(job.folder / "drawing.svg")
        assert len(converted["size"]) == 4 and all(value > 0 for value in converted["size"][:2])
        assert converted["warning"]
        _, plan = prepare(job)
        assert plan["texts"] and all(not item["ocr"] for item in plan["texts"])
        native_ids = sorted({item["item"]["SourceId"] for item in plan["texts"]})
        for item in plan["texts"]:
            assert not item["unsupported"], item
            assert 1 <= item["fontSize"] <= 4000
            assert item["metrics"]["width"] > 0 and item["metrics"]["height"] > 0
            assert abs(item["metrics"]["height"] - item["targetHeight"]) < 0.1
        _, finalized = job.run("finalize", fallbackIds=[])
        assert finalized["hasGeometry"]
        assert finalized["width"] == plan["width"]
        prepare(job)
        _, restored = job.run("finalize", fallbackIds=native_ids)
        svg = ET.parse(restored["path"])
        assert any("中文" in (element.text or "") for element in svg.iter("{http://www.w3.org/2000/svg}text"))
        assert digest(drawing) == original, "The source drawing was modified"
        print("PASS actual DWG conversion, Chinese text, fitted metrics, finalize and native fallback")
    finally:
        job.cleanup()


def test_integrity_and_phase_guards(worker, drawing):
    job = Job(worker)
    try:
        job.run("finalize", expect=False, fallbackIds=[])
        job.run("convert", input=str(drawing), ocr=False)
        prepare(job)
        job.run("finalize", expect=False, fallbackIds=["not-a-native-source-id"])
        job.run("prepare", expect=False, font="unapproved-font", pageWidth=960, pageHeight=540, ocrSelections=[])
        geometry = job.folder / "geometry.svg"
        original = geometry.read_bytes()
        geometry.write_bytes(original + b" ")
        status, _ = job.run("prepare", expect=False, font="Arial", pageWidth=960, pageHeight=540, ocrSelections=[])
        assert "变化" in status["error"]
        geometry.write_bytes(original)
        _, plan = job.run("prepare", font="Arial", pageWidth=960, pageHeight=540, ocrSelections=[])
        assert plan["texts"]
        print("PASS phase ordering, allowed fonts, native-only fallback and immutable-file tamper rejection")
    finally:
        job.cleanup()


def test_cancellation(worker, drawing):
    job = Job(worker)
    marker = Path(str(job.prefix) + ".cancel.flag")
    marker.write_text("1", encoding="ascii")
    status, _ = job.run("convert", expect=False, input=str(drawing), ocr=False)
    assert status["cancelled"]
    assert not job.folder.exists()
    job.cleanup()
    job = Job(worker)
    try:
        request = job.request("convert", input=str(drawing), ocr=True)
        process = subprocess.Popen([str(worker), "--job", str(request)], creationflags=FLAGS)
        deadline = time.monotonic() + 10
        while not job.status("convert").exists():
            assert process.poll() is None, "Worker exited before publishing pending status"
            assert time.monotonic() < deadline
            time.sleep(0.02)
        conflicting = job.request("prepare", font="Arial", pageWidth=960, pageHeight=540, ocrSelections=[])
        duplicate = subprocess.Popen([str(worker), "--job", str(conflicting)], creationflags=FLAGS)
        time.sleep(0.15)
        assert duplicate.poll() is None, "A queued phase should wait briefly for the active phase"
        Path(str(job.prefix) + ".cancel.flag").write_text("1", encoding="ascii")
        assert process.wait(timeout=15) != 0
        assert duplicate.wait(timeout=15) != 0
        queued_status = json.loads(job.status("prepare").read_text(encoding="ascii"))
        assert queued_status["state"] == "done" and queued_status["cancelled"], queued_status
        status = json.loads(job.status("convert").read_text(encoding="utf-8"))
        assert status["cancelled"] and not status["ok"], status
        assert not job.folder.exists(), "Cancelled conversion left a large private folder"
        print("PASS pre-start/running cancellation, phase exclusion, child termination and owned-folder cleanup")
    finally:
        job.cleanup()


def test_phase_lock_handoff(worker, drawing):
    job = Job(worker)
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, wintypes.LPVOID, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    kernel.CreateFileW.restype = wintypes.HANDLE
    kernel.CloseHandle.argtypes = [wintypes.HANDLE]
    handle = None
    try:
        job.run("convert", input=str(drawing), ocr=False)
        handle = kernel.CreateFileW(str(job.prefix) + ".lock", 0xC0000000, 1, None, 3, 0x80, None)
        assert handle != wintypes.HANDLE(-1).value, ctypes.get_last_error()
        request = job.request("prepare", font="Arial", pageWidth=960, pageHeight=540, ocrSelections=[])
        queued = subprocess.Popen([str(worker), "--job", str(request)], creationflags=FLAGS)
        time.sleep(0.3)
        assert queued.poll() is None, "The worker exited instead of waiting for its predecessor to release the lock"
        kernel.CloseHandle(handle)
        handle = None
        assert queued.wait(timeout=15) == 0
        status = json.loads(job.status("prepare").read_text(encoding="ascii"))
        assert status["ok"] and status["state"] == "done", status
        print("PASS phase handoff retries until the predecessor releases its job lock")
    finally:
        if handle is not None:
            kernel.CloseHandle(handle)
        job.cleanup()


def test_ocr_preparation_contract(worker, drawing):
    job = Job(worker)
    try:
        _, converted = job.run("convert", input=str(drawing), ocr=False)
        size = converted["size"]
        page = {
            "width": 1600,
            "height": 1600,
            "viewBox": [size[2], size[3], size[0], size[1]],
            "lines": [{
                "text": "入口轮廓识别",
                "confidence": 0.93,
                "color": "#234567",
                "points": [{"x": 100, "y": 100}, {"x": 500, "y": 100}, {"x": 500, "y": 180}, {"x": 100, "y": 180}],
            }],
        }
        # Inject a known recognition snapshot to test worker layout semantics
        # independently of model confidence and image-dependent OCR output.
        ocr_path = job.folder / "ocr-layout.json"
        ocr_path.write_text(json.dumps(page, ensure_ascii=True), encoding="ascii")
        state_path = job.folder / "state.json"
        state = json.loads(state_path.read_text(encoding="ascii"))
        state["ocr"] = True
        state["hashes"]["ocr-layout.json"] = digest(ocr_path)
        state_path.write_text(json.dumps(state), encoding="ascii")
        _, plan = job.run("prepare", font="Microsoft YaHei", pageWidth=960, pageHeight=540,
                          ocrSelections=[{"index": 0, "text": "核对后的入口文字"}])
        recognized = [item for item in plan["texts"] if item["ocr"]]
        assert len(recognized) == 1 and recognized[0]["item"]["Text"] == "核对后的入口文字"
        assert not recognized[0]["unsupported"] and recognized[0]["metrics"]["height"] > 0
        _, finalized = job.run("finalize", fallbackIds=[])
        assert ET.parse(finalized["path"]).findall(".//{http://www.w3.org/2000/svg}clipPath")
        for selections in [
            [{"index": 0, "text": ""}],
            [{"index": 0, "text": "两行\n文字"}],
            [{"index": 0, "text": "重复"}, {"index": 0, "text": "重复"}],
            [{"index": 2, "text": "越界"}],
            [{"index": index, "text": "超限"} for index in range(501)],
        ]:
            job.run("prepare", expect=False, font="Microsoft YaHei", pageWidth=960, pageHeight=540, ocrSelections=selections)
        print("PASS selected OCR corrected text, fitted plan, clipping and invalid/duplicate/over-limit rejection")
    finally:
        job.cleanup()


def test_request_identity_and_path(worker):
    job = Job(worker)
    request = job.request("convert", input="missing.dwg", ocr=False)
    value = json.loads(request.read_text(encoding="ascii"))
    value["id"] = uuid.uuid4().hex
    request.write_text(json.dumps(value), encoding="ascii")
    process = subprocess.run([str(worker), "--job", str(request)], timeout=15, creationflags=FLAGS)
    status = json.loads(job.status("convert").read_text(encoding="ascii"))
    assert process.returncode != 0 and not status["ok"] and not job.folder.exists()
    job.cleanup()
    with tempfile.TemporaryDirectory(prefix="WpsCadOutside-") as other:
        forbidden = Path(other) / ("DongsidayeWpsCad_" + uuid.uuid4().hex + ".cleanup.json")
        forbidden.write_text("{}", encoding="ascii")
        process = subprocess.run([str(worker), "--job", str(forbidden)], timeout=15, creationflags=FLAGS)
        assert process.returncode != 0 and forbidden.read_text() == "{}"
    print("PASS request identity matching and exact TEMP-root path boundary")


def test_short_temp_alias(worker, drawing):
    job = Job(worker)
    short_temp = Path(os.environ["TEMP"])
    assert short_temp.resolve() == TEMP
    job.prefix = short_temp / job.prefix.name
    job.folder = job.prefix
    try:
        status, converted = job.run("convert", input=str(drawing), ocr=False)
        assert status["output"].startswith(str(job.prefix) + os.sep), status
        assert converted["textCount"] > 0
        _, plan = prepare(job)
        assert plan["texts"]
        status, finalized = job.run("finalize", fallbackIds=[])
        assert finalized["path"].startswith(str(job.prefix) + os.sep), finalized
        print("PASS short TEMP alias acceptance, caller-spelled output paths and cleanup")
    finally:
        job.cleanup()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--worker", type=Path, default=ROOT / "build/cad/WpsCadBridge.exe")
    parser.add_argument("--fixture", type=Path)
    options = parser.parse_args()
    assert options.worker.is_file(), "Run wps_addin/build_cad.ps1 first"
    with tempfile.TemporaryDirectory(prefix="WpsCadWorkerTests-") as work:
        drawing = options.fixture
        if drawing is None:
            fixture_folder = Path(work) / "fixture"
            subprocess.run(["dotnet", "run", "--project", str(REPO / "ppt_addin/cad_converter/smoke/SmokeTest.csproj"), "-c", "Release", "--", str(fixture_folder)], check=True, timeout=120, creationflags=FLAGS, stdout=subprocess.PIPE)
            drawing = next(fixture_folder.glob("*.dwg"))
        test_real_conversion(options.worker.resolve(), drawing.resolve())
        test_integrity_and_phase_guards(options.worker.resolve(), drawing.resolve())
        test_ocr_preparation_contract(options.worker.resolve(), drawing.resolve())
        test_request_identity_and_path(options.worker.resolve())
        test_short_temp_alias(options.worker.resolve(), drawing.resolve())
        test_phase_lock_handoff(options.worker.resolve(), drawing.resolve())
        test_cancellation(options.worker.resolve(), drawing.resolve())


if __name__ == "__main__":
    main()
