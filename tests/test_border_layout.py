"""Issue #5: reference borders only on first generated worksheet.

Approved source sample: 123-Yoklama-dogru.xlsx (sheet one).
Wrong source: 123-Yoklama-yanlis.xlsx.  The source examples are not committed.
"""
from copy import copy
from pathlib import Path
from tempfile import TemporaryDirectory
import subprocess,sys,types
from openpyxl import load_workbook

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
import adaptive_layout as new

BASE='a0de32cf23196551b028601acbfb83d44d2ab4b0'
old=types.ModuleType('_adaptive_layout_v14');old.__file__=str(ROOT/'adaptive_layout.py')
sys.modules[old.__name__]=old
exec(compile(subprocess.check_output(['git','show',BASE+':adaptive_layout.py'],
             cwd=ROOT,text=True),old.__file__,'exec'),old.__dict__)

def side_signature(cell,edge):
    s=getattr(cell.border,edge)
    if s is None or s.style is None:return None
    return (s.style,s.color.type,s.color.indexed if s.color.type=='indexed' else s.color.rgb)

def black_side(cell,edge,kind):
    got=side_signature(cell,edge)
    assert got==(kind,'indexed',64),(cell.coordinate,edge,got)

def identical_sheet(ws,reference):
    assert ws.title==reference.title
    assert set(map(str,ws.merged_cells.ranges))==set(map(str,reference.merged_cells.ranges))
    assert ws.print_area==reference.print_area
    assert ws.print_title_rows==reference.print_title_rows
    assert ws.page_setup.orientation==reference.page_setup.orientation
    assert ws.page_setup.fitToWidth==reference.page_setup.fitToWidth
    assert len(ws.row_breaks.brk)==len(reference.row_breaks.brk)
    assert set(ws.row_dimensions)==set(reference.row_dimensions)
    assert set(ws.column_dimensions)==set(reference.column_dimensions)
    for row in ws.row_dimensions:
        assert ws.row_dimensions[row].height==reference.row_dimensions[row].height
    for col in ws.column_dimensions:
        assert ws.column_dimensions[col].width==reference.column_dimensions[col].width
    assert set(ws._cells)==set(reference._cells)
    for position,cell in ws._cells.items():
        prev=reference._cells[position]
        assert cell.value==prev.value
        for field in ('font','fill','border','alignment','protection'):
            assert copy(getattr(cell,field))==copy(getattr(prev,field)),(position,field)

def verify(weeks,signatures,midterm,final,start,student_count):
    students=[(str(2026000+i),'SENTETİK ÖĞRENCİ '+str(i)) for i in range(student_count)]
    args=(students,'BÖLÜM','DERS','HOCA','BAŞLIK',
          weeks,signatures,midterm,final,15,8,'Arial',58,start)
    with TemporaryDirectory() as d:
        a,b=Path(d)/'old.xlsx',Path(d)/'new.xlsx'
        for module,path in ((old,a),(new,b)):
            wb,_,_=module.build_workbook(*args);wb.save(path);wb.close()
        original,current=(load_workbook(p) for p in (a,b))
        assert original.sheetnames==current.sheetnames
        assert len(original.worksheets)==len(current.worksheets)
        # This is the user-specified guardrail: sheet 2 onward is untouched.
        for ws,previous in zip(current.worksheets[1:],original.worksheets[1:]):
            identical_sheet(ws,previous)
        ws=current.worksheets[0]
        previous=original.worksheets[0]
        assert set(map(str,ws.merged_cells.ranges))==set(map(str,previous.merged_cells.ranges))
        assert ws.print_area==previous.print_area
        assert ws.print_title_rows=='$1:$6'
        assert len(ws.row_breaks.brk)==len(previous.row_breaks.brk)
        last_col,last_row=ws.max_column,ws.max_row
        black_side(ws['A1'],'left','medium')
        black_side(ws['A1'],'top','medium')
        black_side(ws['B1'],'top','medium')
        black_side(ws.cell(1,last_col),'right','medium')
        black_side(ws['A2'],'left','medium')
        black_side(ws['A5'],'left','thin')
        black_side(ws['A5'],'top','thin')
        black_side(ws['B2'],'bottom','thin')
        black_side(ws.cell(6,last_col),'right','thin')
        black_side(ws['A7'],'left','medium')
        black_side(ws.cell(7,last_col),'right','medium')
        black_side(ws.cell(last_row,3),'bottom','medium')
        black_side(ws.cell(last_row,last_col),'bottom','medium')
        if last_col>5 and last_row>8:
            assert copy(ws['D8'].border)==copy(previous['D8'].border)
            assert copy(ws['B8'].border)==copy(previous['B8'].border)
        assert ws['A5'].value=='SIRA'  # This label is inadvertently blank in reference.
        for after,before in zip(current.worksheets,original.worksheets):
            assert set(after._cells)==set(before._cells)
            for location,c in before._cells.items():
                value=after._cells[location]
                assert c.value==value.value
                assert copy(c.font)==copy(value.font)
                assert copy(c.fill)==copy(value.fill)
                assert copy(c.alignment)==copy(value.alignment)
        original.close();current.close()

if __name__=='__main__':
    trials=0
    for weeks,signatures,midterm,final,starts in [
        (16,3,8,16,(1,4,6,9,15)),
        (16,1,8,16,(1,4,9)),
        (20,5,8,20,(1,4,9,15))]:
        for start in starts:
            for n in (4,40,105):
                verify(weeks,signatures,midterm,final,start,n)
                trials+=1
    print('Issue #5 border layout and sheet-two preservation:',trials,'cases passed')
