"""v13 acceptance: independent closed widths, active marks, A4, multi-page merges."""
import sys
from pathlib import Path
from tempfile import TemporaryDirectory
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
import adaptive_layout as a

def verify(weeks,sigs,mid,final,minmark,start,n):
    students=[(str(2026000+i),'ÖĞRENCİ '+str(i)) for i in range(n)]
    with TemporaryDirectory() as temp:
        wb,manifest,_=a.build_workbook(students,'BÖLÜM','DERS','HOCA','BAŞLIK',
                        weeks,sigs,mid,final,minmark,8,'Arial',58,start)
        path=Path(temp)/'out.xlsx'
        wb.save(path);wb.close()
        result=load_workbook(path)
        assert len(result.worksheets)==len(manifest)
        covered=[g['week'] for sh in manifest for g in sh['groups']]
        assert all(x in covered for x in range(start,weeks+1))
        for ws,sh in zip(result.worksheets,manifest):
            groups=sh['groups'];marks=0;closed=[]
            for g in groups:
                old=(g['kind']=='week' and g['week']<start)
                if old:
                    assert g['span']==1
                    assert g['widthMm']==a.CLOSED_WIDTH_MM
                else:
                    assert g['span']==(1 if g['kind']=='exam' else sigs)
                    assert g['widthMm']==sh['markWidthMm']
                for offset in range(g['span']):
                    col=4+marks+offset
                    width=ws.column_dimensions[get_column_letter(col)].width
                    assert width==a.mm_to_excel_chars(g['widthMm'])
                marks+=g['span']
                if old:closed.append(g['week'])
            assert sh['totalWidthMm']<=a.PRINTABLE_MM+.01
            assert sh['markWidthMm']>=minmark-.001
            assert ws.max_column==3+marks
            blocks=[r for r in ws.merged_cells.ranges if r.min_row>=7]
            expected_runs=a.closed_runs([a.Group(g['week'],g['kind'],g['label'],g['span']) for g in groups],start)
            assert len(blocks)==len(expected_runs)*len(sh['pages'])
            for block in blocks:
                assert block.max_row-block.min_row+1==17
                assert ws.cell(block.min_row,block.min_col).fill.fgColor.rgb.endswith(a.EXAM_GRAY)
            assert ws.print_title_rows=='$1:$6'
            assert ws.page_setup.fitToWidth==1
            assert ws.page_setup.orientation=='landscape'
        result.close()

def main():
    previous=a.page_plan(16,3,8,16,15)
    compact=a.page_plan(16,3,8,16,15,a.FIXED_MM,4)
    assert len(compact)<len(previous)
    for start in [1,2,4,8,9,15]:
        for n in [1,40,105]:
            verify(16,3,8,16,15,start,n)
    for start in [2,6,7,14]:
        verify(16,1,6,16,8,start,40)
    for start in [2,8,17]:
        verify(20,5,8,20,20,start,105)
    for start in [2,4,8]:
        verify(10,2,4,10,20,start,17)
    print('v13 Excel and PDF manifest compact layouts: OK')

if __name__=='__main__':
    main()
