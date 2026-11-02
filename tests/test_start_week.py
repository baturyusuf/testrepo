"""Regression tests for start week, including multi-page Excel merges."""
import json
import re
import sys
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory
from openpyxl import load_workbook

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
import adaptive_layout as layout
import web_bridge

CONFIGS=[
    (16,1,8,16,15), (16,2,8,16,15), (16,3,8,16,15),
    (20,5,8,16,15), (10,2,4,9,20), (20,1,6,19,8), (12,1,1,12,15)
]

def students(n):
    return [(str(2026000+i),'ÖĞRENCİ '+str(i)) for i in range(n)]

def workbook(path,config,n,start=None):
    weeks,signatures,midterm,final,minmark=config
    arguments=(students(n),'BÖLÜM','DERS','HOCA','BAŞLIK',
               weeks,signatures,midterm,final,minmark,8,'Arial',58)
    if start is None: wb,manifest,_=layout.build_workbook(*arguments)
    else: wb,manifest,_=layout.build_workbook(*arguments,start)
    wb.save(path)
    wb.close()
    return manifest

def print_config(path,index):
    with zipfile.ZipFile(path) as archive:
        xml=archive.read('xl/worksheets/sheet'+str(index+1)+'.xml').decode()
    return (
      re.findall(r'<col [^>]*>',xml),
      re.findall(r'<row r="(\d+)"[^>]*?ht="([\d.]+)"',xml),
      re.findall(r'<rowBreaks.*?</rowBreaks>',xml),
      re.findall(r'<(?:pageSetup|pageMargins|printOptions|sheetPr|pageSetUpPr)[^>]*>',xml)
    )

def verify(config,n):
    weeks,signatures,midterm,final,_=config
    total=0
    with TemporaryDirectory() as directory:
        root=Path(directory)
        original=root/'first.xlsx'
        all_sections=workbook(original,config,n,1)
        for start in range(2,layout.last_signature_week(weeks,midterm,final)+1):
            output=root/'modified.xlsx'
            sections=workbook(output,config,n,start)
            kept=[(i,s) for i,s in enumerate(all_sections) if
                  any(g['kind']=='week' and g['week']>=start for g in s['groups'])]
            assert [s['name'] for s in sections]==[s['name'] for _,s in kept]
            assert len(sections)==len(kept)>=1
            opened=load_workbook(output)
            assert len(opened.worksheets)==len(kept)
            for index,(ws,section,(base_index,base)) in enumerate(zip(opened.worksheets,sections,kept)):
                assert print_config(output,index)==print_config(original,base_index)
                assert section['markWidthMm']==base['markWidthMm']
                assert section['startWeek']==start and section['totalSections']==len(all_sections)
                blocks=sorted((r.min_col,r.max_col,r.min_row,r.max_row) for r in ws.merged_cells.ranges
                              if r.min_row>layout.HEADER_ROWS)
                expected=[]
                col=4
                for g in section['groups']:
                    closed=g['kind']=='week' and g['week']<start
                    assert bool(ws.cell(5,col).font.strike)==closed
                    if g['kind']=='exam':
                        assert ws.cell(7,col).value is None
                        assert not ws.cell(5,col).font.strike
                    if closed:
                        if expected and expected[-1][1]+1==col:
                            expected[-1][1]=col+g['span']-1
                        else:
                            expected.append([col,col+g['span']-1])
                        if g['span']>1:
                            assert all(ws.cell(6,col+k).font.strike for k in range(g['span']))
                    col+=g['span']
                wanted=sorted((a,b,7+page*17,7+page*17+16)
                              for a,b in expected for page in range(len(section['pages'])))
                assert blocks==wanted,(config,n,start,blocks,wanted)
                for a,b,top,_ in blocks:
                    c=ws.cell(top,a)
                    assert c.value.startswith(str(start)+'. HAFTADAN ÖNCEKİ HAFTALAR')
                    assert c.value.endswith('BU LİSTEDE İMZALANMAZ')
                    assert c.fill.fgColor.rgb.endswith(layout.EXAM_GRAY)
                    assert c.alignment.textRotation==(90 if (b-a+1)*section['markWidthMm']<32 else 0)
                assert ws.cell(7,1).value==1
                assert ws.cell(7,3).value=='ÖĞRENCİ 0'
            opened.close()
            total+=1
    return total

def verify_first_week():
    with TemporaryDirectory() as directory:
        root=Path(directory)
        for cfg in CONFIGS:
            a=root/'a.xlsx';b=root/'b.xlsx'
            workbook(a,cfg,40,None)
            workbook(b,cfg,40,1)
            with zipfile.ZipFile(a) as first,zipfile.ZipFile(b) as second:
                for name in first.namelist():
                    if name!='docProps/core.xml':
                        assert first.read(name)==second.read(name),name

def verify_validation():
    data={'weeks':16,'signatures':1,'midterm':8,'final':16,'course':'D',
          'instructor':'H','academicYear':'2026-2027','semester':'Güz','files':[]}
    with TemporaryDirectory() as directory:
        web_bridge.WORK=Path(directory)
        assert json.loads(web_bridge.generate(json.dumps(data)))['configuration']['startWeek']==1
    assert layout.validate_start_week('4',16,8,16)==4
    assert layout.validate_start_week(16,16,8,15)==16
    for bad in (0,-1,16,17,'x','4.5',None,True,2.5):
        try:
            web_bridge.generate(json.dumps({**data,'startWeek':bad}))
        except ValueError: pass
        else: raise AssertionError('Invalid startWeek accepted: '+repr(bad))


def verify_report_examples():
    """Independent expected Excel column ranges from the supplied acceptance report."""
    default=layout.page_plan(16,1,8,16,15)
    assert [[g.week for g in section] for section in default]==[list(range(1,9)),list(range(9,17))]
    assert layout.closed_runs(default[0],2)==[(4,4)]
    assert layout.closed_runs(default[0],3)==[(4,5)]
    assert layout.closed_runs(default[0],4)==[(4,6)]
    assert layout.closed_runs(default[1],10)==[(4,4)]
    assert layout.closed_runs(default[1],15)==[(4,9)]
    assert not layout.has_open_week(default[0],8)
    triple=layout.page_plan(16,3,8,16,15)
    assert [g[0].week for g in triple]==[1,5,9,13]
    assert layout.closed_runs(triple[0],4)==[(4,12)]
    assert layout.closed_runs(triple[1],6)==[(4,6)]
    assert not layout.has_open_week(triple[1],8)
    with_exam=layout.page_plan(16,1,6,16,15)
    assert layout.closed_runs(with_exam[0],8)==[(4,8),(10,10)]
    assert with_exam[0][5].kind=='exam'

if __name__=='__main__':
    checked=sum(verify(c,n) for c in CONFIGS for n in (1,40,105))
    verify_first_week()
    verify_report_examples()
    verify_validation()
    print('Start week Excel regression: '+str(checked)+' configurations passed')
