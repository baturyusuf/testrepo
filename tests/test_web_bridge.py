"""Integration smoke tests for the three attendance Excel templates."""
from pathlib import Path
from tempfile import TemporaryDirectory
import json
import shutil
import sys

from openpyxl import Workbook, load_workbook

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))
import web_bridge  # noqa: E402

with TemporaryDirectory() as directory:
    work = Path(directory)
    (work / "templates").mkdir()
    (work / "uploads" / "1").mkdir(parents=True)
    for template in (PROJECT / "templates").glob("*.xlsx"):
        shutil.copy2(template, work / "templates" / template.name)

    incoming = Workbook()
    sheet = incoming.active
    sheet.append(["Öğrenci No_123", "Adı_123", "Soyadı_123", "Program_123"])
    sheet.append([123456, "Ayşe", "Yılmaz", "Bilgisayar Mühendisliği"])
    sheet.append([123457, "Özgür", "Demir", "Bilgisayar Mühendisliği"])
    incoming.save(work / "uploads" / "1" / "205-bb.xlsx")
    web_bridge.WORK = work

    for signatures in (1, 2, 3):
        for semester in ("Güz", "Bahar"):
            payload = {
                "signatures": signatures,
                "course": "BLG103 - Programlama I",
                "instructor": "Dr. Öğr. Üyesi Deneme",
                "academicYear": "2027-2028",
                "semester": semester,
                "files": [{"name": "205-bb.xlsx"}],
            }
            result = json.loads(web_bridge.generate(json.dumps(payload)))
            assert len(result["results"]) == 1, result
            output = result["results"][0]
            workbook = load_workbook(output["file"])
            for ws in workbook:
                assert "2027-2028" in ws["A1"].value
                assert semester.upper() in ws["A1"].value
            assert all("2027-2028" in info["title"] for info in output["sheets"])
            assert len(output["sheets"]) == (1 if signatures == 1 else 2)
            workbook.close()

    invalid = dict(payload, academicYear="2026-2028")
    try:
        web_bridge.generate(json.dumps(invalid))
    except ValueError as exc:
        assert "Eğitim öğretim yılı" in str(exc)
    else:
        raise AssertionError("Invalid academic year was accepted")

print("Excel and PDF layout integration: OK (three templates, two semesters)")
