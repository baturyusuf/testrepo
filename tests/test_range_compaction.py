"""Issue #8: Python Excel/manifest verifies compact closed weeks."""
import sys
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
import adaptive_layout as a
from openpyxl.utils import get_column_letter
def check(start,mid=8,n=40):
 students=[(str(1000+i),f'ÖĞRENCİ {i}') for i in range(n)]
 wb,manifest,_=a.build_workbook(students,'BÖLÜM','DERS','HOCA','BAŞLIK',
                                16,2,mid,16,10,8,'Arial',58,start)
 assert len(wb.worksheets)==len(manifest)
 for ws,sheet in zip(wb.worksheets,manifest):
  assert sheet['totalWidthMm']<=289.01
  assert sheet['markWidthMm']>=10
  col=4
  for g in sheet['groups']:
   assert ws.column_dimensions[get_column_letter(col)].width==a.mm_to_excel_chars(g['widthMm'])
   if g['kind']=='closed':
    assert g['span']==1 and g['widthMm']==26
    assert ws.cell(5,col).value==g['label']
    assert ws.cell(7,col).value==(
       f"{g['week']}-{g['endWeek']} arasındaki haftalar\n"
       "bu belgede imzalanamaz")
    assert ws.cell(7,col).alignment.textRotation==0
   col+=g['span']
 wb.close()
 return [g for part in a.page_plan(16,2,mid,16,10,a.FIXED_MM,start) for g in part]
nine=check(9,n=105)
assert any(g.kind=='closed' and g.week==1 and g.end_week==7 for g in nine)
assert any(g.kind=='exam' and g.week==8 for g in nine)
assert not any(g.kind=='closed' for g in check(4))
assert any(g.kind=='closed' and g.week==1 and g.end_week==4 for g in check(5))
split=check(9,4)
assert any(g.kind=='closed' and g.week==5 and g.end_week==8 for g in split)
assert any(g.kind=='exam' and g.week==4 for g in split)
print('Issue #8 Excel/manifest compact blocks: OK')
