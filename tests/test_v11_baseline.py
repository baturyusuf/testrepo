"""Compare v12 week-1 workbook XML against unmodified v11 code, not merely v12 defaults."""
import subprocess
import sys
import types
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
import adaptive_layout as latest
BASE='77e6add481767829bc76f6e4564ac214c0be4928'
old_source=subprocess.check_output(['git','show',BASE+':adaptive_layout.py'],cwd=ROOT,text=True)
old=types.ModuleType('_adaptive_layout_v11')
old.__file__=str(ROOT/'adaptive_layout.py')
sys.modules[old.__name__]=old
exec(compile(old_source,old.__file__,'exec'),old.__dict__)
students=[(str(20260000+n),'ÖĞRENCİ '+str(n)) for n in range(40)]
for weeks,signatures,midterm,final,minimum in [(16,1,8,16,15),(16,3,8,16,15),(20,5,8,16,8)]:
    args=(students,'BÖLÜM','DERS','HOCA','BAŞLIK',weeks,signatures,midterm,final,minimum,8,'Arial',58)
    with TemporaryDirectory() as tmp:
        a,b=Path(tmp)/'v11.xlsx',Path(tmp)/'v12.xlsx'
        for module,path in [(old,a),(latest,b)]:
            wb,_,_=module.build_workbook(*args)
            wb.save(path)
            wb.close()
        with zipfile.ZipFile(a) as previous,zipfile.ZipFile(b) as current:
            assert set(previous.namelist())==set(current.namelist())
            for name in previous.namelist():
                if name!='docProps/core.xml':
                    assert previous.read(name)==current.read(name), (name,weeks,signatures)
print('v11 → v12 first-week workbook XML: identical')
