"""v11 comparison: preserve content/geometry; v15 changes sheet-one border styles."""
from copy import copy
from pathlib import Path
from tempfile import TemporaryDirectory
import subprocess,sys,types
from openpyxl import load_workbook

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
import adaptive_layout as latest
BASE='77e6add481767829bc76f6e4564ac214c0be4928'
old=types.ModuleType('_adaptive_layout_v11');old.__file__=str(ROOT/'adaptive_layout.py')
sys.modules[old.__name__]=old
exec(compile(subprocess.check_output(['git','show',BASE+':adaptive_layout.py'],
            cwd=ROOT,text=True),old.__file__,'exec'),old.__dict__)
students=[(str(20260000+n),'ÖĞRENCİ '+str(n)) for n in range(40)]
for weeks,signatures,midterm,final,minimum in [
    (16,1,8,16,15),(16,3,8,16,15),(20,5,8,16,8)]:
    args=(students,'BÖLÜM','DERS','HOCA','BAŞLIK',
          weeks,signatures,midterm,final,minimum,8,'Arial',58)
    with TemporaryDirectory() as d:
        paths=[Path(d)/'v11.xlsx',Path(d)/'v15.xlsx']
        for m,p in zip([old,latest],paths):
            wb,_,_=m.build_workbook(*args);wb.save(p);wb.close()
        before,after=(load_workbook(p) for p in paths)
        assert before.sheetnames==after.sheetnames
        for idx,(ws,updated) in enumerate(zip(before.worksheets,after.worksheets)):
            assert set(map(str,ws.merged_cells.ranges))==set(map(str,updated.merged_cells.ranges))
            assert ws.print_area==updated.print_area
            assert ws.print_title_rows==updated.print_title_rows
            assert ws.page_setup.orientation==updated.page_setup.orientation
            assert ws.page_setup.fitToWidth==updated.page_setup.fitToWidth
            assert set(ws.row_dimensions)==set(updated.row_dimensions)
            assert set(ws.column_dimensions)==set(updated.column_dimensions)
            for r in ws.row_dimensions:
                assert ws.row_dimensions[r].height==updated.row_dimensions[r].height
            for c in ws.column_dimensions:
                assert ws.column_dimensions[c].width==updated.column_dimensions[c].width
            assert set(ws._cells)==set(updated._cells)
            for address,c in ws._cells.items():
                got=updated._cells[address]
                assert got.value==c.value
                for field in ('font','fill','alignment','protection'):
                    assert copy(getattr(c,field))==copy(getattr(got,field))
                if idx:
                    assert copy(c.border)==copy(got.border)
        before.close();after.close()
print('v11 → v15: content, geometry and subsequent-sheet styles preserved')
