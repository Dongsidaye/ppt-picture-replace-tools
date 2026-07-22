"""Inject a RibbonX customUI14 part without corrupting empty PPAM ZIP entries."""

from __future__ import annotations

import os
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
UI_REL = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
UI_PART = "customUI/customUI14.xml"


def clone_info(source: zipfile.ZipInfo) -> zipfile.ZipInfo:
    target = zipfile.ZipInfo(source.filename, source.date_time)
    target.comment = source.comment
    target.extra = source.extra
    target.create_system = source.create_system
    target.create_version = source.create_version
    target.extract_version = source.extract_version
    target.internal_attr = source.internal_attr
    target.external_attr = source.external_attr
    target.volume = source.volume
    target.compress_type = source.compress_type
    return target


def inject(ppam_path: Path, custom_ui_path: Path) -> None:
    ui_bytes = custom_ui_path.read_bytes()
    ET.fromstring(ui_bytes)  # fail fast on malformed RibbonX

    with zipfile.ZipFile(ppam_path, "r") as source:
        entries = [(info, source.read(info.filename)) for info in source.infolist()]

    rel_name = "_rels/.rels"
    rel_bytes = next((data for info, data in entries if info.filename == rel_name), None)
    if rel_bytes is None:
        raise RuntimeError("PPAM is missing _rels/.rels")

    ET.register_namespace("", REL_NS)
    root = ET.fromstring(rel_bytes)
    for rel in list(root):
        if rel.get("Type") == UI_REL:
            root.remove(rel)
    used_ids = {rel.get("Id") for rel in root}
    rel_id = "rId2"
    counter = 2
    while rel_id in used_ids:
        counter += 1
        rel_id = f"rId{counter}"
    ET.SubElement(root, f"{{{REL_NS}}}Relationship", {
        "Id": rel_id,
        "Type": UI_REL,
        "Target": UI_PART,
    })
    new_rels = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    fd, temp_name = tempfile.mkstemp(prefix=ppam_path.stem + "-", suffix=".ppam", dir=ppam_path.parent)
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        with zipfile.ZipFile(temp_path, "w", allowZip64=False) as target:
            for info, data in entries:
                if info.filename in {rel_name, UI_PART, "customUI/customUI.xml"}:
                    continue
                cloned = clone_info(info)
                target.writestr(cloned, data, compress_type=info.compress_type, compresslevel=6)
            target.writestr(rel_name, new_rels, compress_type=zipfile.ZIP_DEFLATED, compresslevel=6)
            target.writestr(UI_PART, ui_bytes, compress_type=zipfile.ZIP_STORED)

        with zipfile.ZipFile(temp_path, "r") as check:
            required = {
                "[Content_Types].xml",
                rel_name,
                "ppt/presentation.xml",
                "ppt/vbaProject.bin",
                UI_PART,
            }
            missing = sorted(required.difference(check.namelist()))
            if missing:
                raise RuntimeError("missing package parts: " + ", ".join(missing))
            presentation = check.getinfo("ppt/presentation.xml")
            if presentation.file_size == 0 and presentation.compress_type != zipfile.ZIP_DEFLATED:
                raise RuntimeError("empty presentation.xml lost its DEFLATE method")
            if check.getinfo("ppt/vbaProject.bin").file_size < 1024:
                raise RuntimeError("vbaProject.bin is empty or truncated")
            if check.getinfo(UI_PART).compress_type != zipfile.ZIP_STORED:
                raise RuntimeError("customUI14.xml must be stored without compression")
            ET.fromstring(check.read(rel_name))
            ET.fromstring(check.read(UI_PART))

        os.replace(temp_path, ppam_path)
    finally:
        if temp_path.exists():
            temp_path.unlink()


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit("usage: inject_customui.py <addin.ppam> <customUI.xml>")
    inject(Path(sys.argv[1]).resolve(), Path(sys.argv[2]).resolve())
    print("customUI14 injected; ZIP invariants passed")
