"""v16: approved border layout and blank serial heading on every generated worksheet.

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

def verify(weeks, signatures, midterm, final, start, student_count):
    """Compare all sheets with v14 for geometry and check v15 borders everywhere."""
    students=[(str(2026000+i),'SENTETİK ÖĞRENCİ '+str(i)) for i in range(student_count)]
    args=(students,'BÖLÜM','DERS','HOCA','BAŞLIK',
          weeks,signatures,midterm,final,15,8,'Arial',58,start)
    with TemporaryDirectory() as d:
        before,after=Path(d)/'v14.xlsx',Path(d)/'v16.xlsx'
        for module,path in ((old,before),(new,after)):
            wb,_,_=module.build_workbook(*args);wb.save(path);wb.close()
        w_old,w_new=load_workbook(before),load_workbook(after)
        assert w_old.sheetnames==w_new.sheetnames
        assert len(w_old.worksheets)==len(w_new.worksheets)
        for ws,original in zip(w_new.worksheets,w_old.worksheets):
            assert ws.print_area==original.print_area
            assert ws.print_title_rows==original.print_title_rows=='$1:$6'
            assert ws.page_setup.orientation==original.page_setup.orientation=='landscape'
            assert ws.page_setup.fitToWidth==original.page_setup.fitToWidth==1
            assert ws.page_setup.paperSize==original.page_setup.paperSize
            assert len(ws.row_breaks.brk)==len(original.row_breaks.brk)
            assert set(map(str,ws.merged_cells.ranges))==set(
                map(str,original.merged_cells.ranges))
            assert set(ws.row_dimensions)==set(original.row_dimensions)
            assert set(ws.column_dimensions)==set(original.column_dimensions)
            for r in ws.row_dimensions:
                assert ws.row_dimensions[r].height==original.row_dimensions[r].height
            for col in ws.column_dimensions:
                assert ws.column_dimensions[col].width==original.column_dimensions[col].width
            assert set(ws._cells)==set(original._cells)
            for key,cell in ws._cells.items():
                prior=original._cells[key]
                assert cell.value==(None if key==(5,1) else prior.value)
                for field in ('font','fill','alignment','protection'):
                    assert copy(getattr(cell,field))==copy(getattr(prior,field)),(
                        ws.title,key,field)
                assert cell.number_format==prior.number_format

            assert ws['A5'].value is None
            assert ws['B5'].value=='ÖĞRENCİ NO'
            assert ws['C5'].value=='AD SOYAD'
            assert any(str(m)=='A5:A6' for m in ws.merged_cells.ranges)
            assert ws['A7'].value==1
            assert copy(ws['A5'].fill)==copy(original['A5'].fill)
            last_col,last_row=ws.max_column,ws.max_row
            black_side(ws['A1'],'left','medium')
            black_side(ws['A1'],'top','medium')
            black_side(ws.cell(1,last_col),'right','medium')
            black_side(ws['B1'],'top','medium')
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
                assert copy(ws['D8'].border)==copy(original['D8'].border)
                assert copy(ws['B8'].border)==copy(original['B8'].border)
        count=len(w_new.worksheets)
        w_old.close();w_new.close()
        return count

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
    print('v16 all-sheet approved borders and blank SIRA heading:',trials,'cases passed')
